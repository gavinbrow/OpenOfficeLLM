import fs from 'node:fs'
import path from 'node:path'
import forge from 'node-forge'
import { resolveCertDir, ensureDirs } from './paths.js'
import { logger } from './logging.js'

const { pki } = forge

/** Exported because the macOS trust path searches the keychain by common name.
 *  A literal copied into trust.ts would drift the first time this changes, and
 *  the symptom would be an installed CA reported as untrusted. */
export const CA_COMMON_NAME = 'OpenOfficeLLM Local CA'

const CA_SUBJECT = [
  { name: 'commonName', value: CA_COMMON_NAME },
  { name: 'organizationName', value: 'OpenOfficeLLM' },
]

const LEAF_SUBJECT = [
  { name: 'commonName', value: 'localhost' },
  { name: 'organizationName', value: 'OpenOfficeLLM' },
]

const CA_CRT = 'ca.crt'
const CA_KEY = 'ca.key'
const LEAF_CRT = 'server.crt'
const LEAF_KEY = 'server.key'
const THUMBPRINT_FILE = 'ca.thumbprint.txt'

const RENEWAL_WINDOW_DAYS = 30

function serial(): string {
  return '00' + forge.util.bytesToHex(forge.random.getBytesSync(16))
}

function yearsFromNow(n: number): Date {
  const d = new Date()
  d.setFullYear(d.getFullYear() + n)
  return d
}

function backdated(): Date {
  const d = new Date()
  d.setMinutes(d.getMinutes() - 10)
  return d
}

interface GeneratedMaterial {
  caCert: forge.pki.Certificate
  caKey: forge.pki.PrivateKey
  leafCert: forge.pki.Certificate
  leafKey: forge.pki.PrivateKey
  thumbprint: string
}

function generateMaterial(): GeneratedMaterial {
  logger.info({ msg: 'generating CA keypair (2048-bit RSA)' })
  const caKeys = pki.rsa.generateKeyPair(2048)
  const ca = pki.createCertificate()
  ca.publicKey = caKeys.publicKey
  ca.serialNumber = serial()
  ca.validity.notBefore = backdated()
  ca.validity.notAfter = yearsFromNow(10)
  ca.setSubject(CA_SUBJECT)
  ca.setIssuer(CA_SUBJECT)
  ca.setExtensions([
    { name: 'basicConstraints', cA: true, critical: true, pathLenConstraint: 0 },
    {
      name: 'keyUsage',
      critical: true,
      keyCertSign: true,
      cRLSign: true,
      digitalSignature: true,
    },
    { name: 'subjectKeyIdentifier' },
  ])
  ca.sign(caKeys.privateKey, forge.md.sha256.create())

  logger.info({ msg: 'generating leaf keypair (2048-bit RSA)' })
  const leafKeys = pki.rsa.generateKeyPair(2048)
  const leaf = pki.createCertificate()
  leaf.publicKey = leafKeys.publicKey
  leaf.serialNumber = serial()
  leaf.validity.notBefore = backdated()
  leaf.validity.notAfter = yearsFromNow(2)
  leaf.setSubject(LEAF_SUBJECT)
  leaf.setIssuer(CA_SUBJECT)
  leaf.setExtensions([
    { name: 'basicConstraints', cA: false, critical: true },
    {
      name: 'keyUsage',
      critical: true,
      digitalSignature: true,
      keyEncipherment: true,
    },
    { name: 'extKeyUsage', serverAuth: true },
    {
      name: 'subjectAltName',
      altNames: [
        { type: 2, value: 'localhost' },
        { type: 7, ip: '127.0.0.1' },
        { type: 7, ip: '::1' },
      ],
    },
    { name: 'subjectKeyIdentifier' },
    {
      name: 'authorityKeyIdentifier',
      keyIdentifier: ca.generateSubjectKeyIdentifier().getBytes(),
    },
  ])
  leaf.sign(caKeys.privateKey, forge.md.sha256.create())

  const thumbprint = computeThumbprint(ca)
  return {
    caCert: ca,
    caKey: caKeys.privateKey,
    leafCert: leaf,
    leafKey: leafKeys.privateKey,
    thumbprint,
  }
}

export function computeThumbprint(cert: forge.pki.Certificate): string {
  const md = forge.md.sha1.create()
  md.update(forge.asn1.toDer(pki.certificateToAsn1(cert)).getBytes())
  return md.digest().toHex().toUpperCase()
}

function certDir(): string {
  return resolveCertDir()
}

function writeMaterial(m: GeneratedMaterial): void {
  ensureDirs()
  const dir = certDir()
  const files: Record<string, { content: string; secret: boolean }> = {
    [CA_CRT]: { content: pki.certificateToPem(m.caCert), secret: false },
    [CA_KEY]: { content: pki.privateKeyToPem(m.caKey), secret: true },
    [LEAF_CRT]: { content: pki.certificateToPem(m.leafCert), secret: false },
    [LEAF_KEY]: { content: pki.privateKeyToPem(m.leafKey), secret: true },
    [THUMBPRINT_FILE]: { content: m.thumbprint, secret: false },
  }
  for (const [name, { content, secret }] of Object.entries(files)) {
    const filePath = path.join(dir, name)
    // Private keys must be owner-only readable. Certs and the thumbprint are
    // public. Default umask typically yields 0644, which would expose the CA
    // private key (root of trust for the loopback cert) to other users on
    // shared machines.
    fs.writeFileSync(filePath, content, { mode: secret ? 0o600 : 0o644 })
    // chmod after write in case the file already existed with looser perms —
    // writeFileSync doesn't change the mode of an existing file.
    if (secret) fs.chmodSync(filePath, 0o600)
  }
}

function readIfExists(name: string): string | null {
  try {
    return fs.readFileSync(path.join(certDir(), name), 'utf8')
  } catch {
    return null
  }
}

function leafNeedsRenewal(leaf: forge.pki.Certificate): boolean {
  const notAfter = leaf.validity.notAfter
  const msLeft = notAfter.getTime() - Date.now()
  return msLeft < RENEWAL_WINDOW_DAYS * 24 * 60 * 60 * 1000
}

export interface CertMaterial {
  caPem: string
  leafCertPem: string
  leafKeyPem: string
  thumbprint: string
  renewed: boolean
}

export function loadCertMaterial(): CertMaterial {
  ensureDirs()
  const caPem = readIfExists(CA_CRT)
  const caKeyPem = readIfExists(CA_KEY)
  const leafPem = readIfExists(LEAF_CRT)
  const leafKeyPem = readIfExists(LEAF_KEY)
  let thumbprint = readIfExists(THUMBPRINT_FILE)

  if (caPem && caKeyPem && leafPem && leafKeyPem) {
    try {
      const leaf = pki.certificateFromPem(leafPem)
      const ca = pki.certificateFromPem(caPem)
      if (!thumbprint) {
        thumbprint = computeThumbprint(ca)
        try {
          fs.writeFileSync(path.join(certDir(), THUMBPRINT_FILE), thumbprint)
        } catch {
          // ignore
        }
      }
      if (leafNeedsRenewal(leaf)) {
        logger.info({
          msg: 'leaf cert expiring soon, regenerating',
          notAfter: leaf.validity.notAfter.toISOString(),
        })
        const caKey = pki.privateKeyFromPem(caKeyPem)
        const leafKey = pki.rsa.generateKeyPair(2048)
        const newLeaf = pki.createCertificate()
        newLeaf.publicKey = leafKey.publicKey
        newLeaf.serialNumber = serial()
        newLeaf.validity.notBefore = backdated()
        newLeaf.validity.notAfter = yearsFromNow(2)
        newLeaf.setSubject(LEAF_SUBJECT)
        newLeaf.setIssuer(CA_SUBJECT)
        newLeaf.setExtensions([
          { name: 'basicConstraints', cA: false, critical: true },
          {
            name: 'keyUsage',
            critical: true,
            digitalSignature: true,
            keyEncipherment: true,
          },
          { name: 'extKeyUsage', serverAuth: true },
          {
            name: 'subjectAltName',
            altNames: [
              { type: 2, value: 'localhost' },
              { type: 7, ip: '127.0.0.1' },
              { type: 7, ip: '::1' },
            ],
          },
          { name: 'subjectKeyIdentifier' },
          {
            name: 'authorityKeyIdentifier',
            keyIdentifier: ca.generateSubjectKeyIdentifier().getBytes(),
          },
        ])
        newLeaf.sign(caKey, forge.md.sha256.create())
        fs.writeFileSync(path.join(certDir(), LEAF_CRT), pki.certificateToPem(newLeaf))
        fs.writeFileSync(path.join(certDir(), LEAF_KEY), pki.privateKeyToPem(leafKey.privateKey))
        logger.info({ msg: 'leaf cert regenerated' })
        return {
          caPem,
          leafCertPem: pki.certificateToPem(newLeaf),
          leafKeyPem: pki.privateKeyToPem(leafKey.privateKey),
          thumbprint,
          renewed: true,
        }
      }
      return { caPem, leafCertPem: leafPem, leafKeyPem, thumbprint, renewed: false }
    } catch (e) {
      logger.warn({ msg: 'existing cert unreadable, regenerating', error: String(e) })
    }
  }

  const m = generateMaterial()
  writeMaterial(m)
  return {
    caPem: pki.certificateToPem(m.caCert),
    leafCertPem: pki.certificateToPem(m.leafCert),
    leafKeyPem: pki.privateKeyToPem(m.leafKey),
    thumbprint: m.thumbprint,
    renewed: false,
  }
}

export function getCaThumbprint(): string | null {
  const t = readIfExists(THUMBPRINT_FILE)
  if (t) return t.trim()
  const caPem = readIfExists(CA_CRT)
  if (!caPem) return null
  try {
    const ca = pki.certificateFromPem(caPem)
    const tp = computeThumbprint(ca)
    try {
      fs.writeFileSync(path.join(certDir(), THUMBPRINT_FILE), tp)
    } catch {
      // ignore
    }
    return tp
  } catch {
    return null
  }
}
