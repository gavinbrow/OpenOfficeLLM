// Phase 0 spike — generate a local CA + leaf cert for https://127.0.0.1:7317
//
// Chromium ignores CN entirely, so the leaf MUST carry SANs. For an IP host it
// must be an *IP* SAN (type 7), not a DNS SAN — a DNS SAN of "127.0.0.1" is
// silently ignored and you get ERR_CERT_COMMON_NAME_INVALID with no clue why.
//
// Chrome's 398-day max validity applies only to certs chaining to a *public*
// root, so the 2-year leaf below is fine for a locally-installed CA.

import forge from 'node-forge'
import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const outDir = join(here, 'certs')
mkdirSync(outDir, { recursive: true })

const { pki } = forge

// Serial must be a positive integer. Leading '00' keeps the high bit clear so
// it doesn't get read as negative.
const serial = () => '00' + forge.util.bytesToHex(forge.random.getBytesSync(16))

function years(n) {
  const d = new Date()
  d.setFullYear(d.getFullYear() + n)
  return d
}

// Backdate slightly so a skewed clock on the machine doesn't reject a
// just-issued cert.
function notBefore() {
  const d = new Date()
  d.setMinutes(d.getMinutes() - 10)
  return d
}

console.log('generating CA keypair (2048-bit)...')
const caKeys = pki.rsa.generateKeyPair(2048)
const ca = pki.createCertificate()
ca.publicKey = caKeys.publicKey
ca.serialNumber = serial()
ca.validity.notBefore = notBefore()
ca.validity.notAfter = years(10)

const caSubject = [
  { name: 'commonName', value: 'OpenOfficeLLM Spike Local CA' },
  { name: 'organizationName', value: 'OpenOfficeLLM' },
]
ca.setSubject(caSubject)
ca.setIssuer(caSubject)
ca.setExtensions([
  { name: 'basicConstraints', cA: true, critical: true, pathLenConstraint: 0 },
  { name: 'keyUsage', critical: true, keyCertSign: true, cRLSign: true, digitalSignature: true },
  { name: 'subjectKeyIdentifier' },
])
ca.sign(caKeys.privateKey, forge.md.sha256.create())

console.log('generating leaf keypair (2048-bit)...')
const leafKeys = pki.rsa.generateKeyPair(2048)
const leaf = pki.createCertificate()
leaf.publicKey = leafKeys.publicKey
leaf.serialNumber = serial()
leaf.validity.notBefore = notBefore()
leaf.validity.notAfter = years(2)

leaf.setSubject([
  { name: 'commonName', value: 'localhost' },
  { name: 'organizationName', value: 'OpenOfficeLLM' },
])
leaf.setIssuer(caSubject)
leaf.setExtensions([
  { name: 'basicConstraints', cA: false, critical: true },
  { name: 'keyUsage', critical: true, digitalSignature: true, keyEncipherment: true },
  { name: 'extKeyUsage', serverAuth: true },
  {
    name: 'subjectAltName',
    altNames: [
      { type: 2, value: 'localhost' }, // 2 = dNSName
      { type: 7, ip: '127.0.0.1' }, // 7 = iPAddress
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

const files = {
  // .crt in PEM form is what PowerShell's Import-Certificate expects.
  'ca.crt': pki.certificateToPem(ca),
  'ca.key': pki.privateKeyToPem(caKeys.privateKey),
  'server.crt': pki.certificateToPem(leaf),
  'server.key': pki.privateKeyToPem(leafKeys.privateKey),
}

for (const [name, content] of Object.entries(files)) {
  writeFileSync(join(outDir, name), content)
  console.log('  wrote', join('certs', name))
}

const md = forge.md.sha1.create()
md.update(forge.asn1.toDer(pki.certificateToAsn1(ca)).getBytes())
const thumbprint = md.digest().toHex().toUpperCase()
writeFileSync(join(outDir, 'ca.thumbprint.txt'), thumbprint)

console.log('\nCA SHA1 thumbprint:', thumbprint)
console.log('(this is what Cert:\\CurrentUser\\Root indexes on — never match by subject name)')
