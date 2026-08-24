import { getSecret } from '../secrets.js'

export function providerApiKey(providerId: string): string | null {
  return getSecret(providerId)
}
