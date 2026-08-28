// Attachment upload helpers shared by the chat drop zone (ChatPanel) and the
// paperclip button (Composer). Both turn a dropped/picked file into a context
// chip, so the store-mutation logic lives here once instead of being copied.
//
// Upload errors are surfaced through the toast system (useUiStore.toast) so
// the user actually sees that their drag failed, rather than the chip
// silently never appearing. The console.error stays for debugging.

import { uploadAttachment } from '../api/client'
import type { AttachmentUploadResult } from '../api/client'
import { useContextStore } from '../store/contextStore'
import { useUiStore } from '../store/uiStore'

/** Convert an upload result into a ContextItem and add it to the context store.
 *  The attachment id doubles as the chip id so removeAttachment() (which
 *  deletes the chip and the host blob in one step) can be keyed off the same
 *  value the chip renders. */
export function addAttachmentItem(result: AttachmentUploadResult): void {
  const store = useContextStore.getState()
  const kind = result.kind === 'image' ? 'image-attachment' : 'text-attachment'
  // Thumbnails are deferred: attachmentUrl(id) needs a bearer token to fetch,
  // and a plain <img src={...}> won't carry it. Fetching the blob, building an
  // object URL, and revoking it on remove is doable but not worth the v1
  // complexity — a file-type icon reads clearly enough. Revisit when image
  // previews matter.
  store.addAttachment({
    id: result.id,
    label: result.fileName,
    scope: 'none',
    tokenEstimate: result.tokenEstimate,
    context: {
      host: 'none',
      scope: 'none',
      // The extracted text is not carried on the ContextItem. The host stores
      // it on the attachment meta and folds it into the system prompt at
      // buildAttachmentPayload time via getMeta(); toDocumentContext() excludes
      // attachment items anyway, so anything set here is dead data that never
      // reaches the model. Leave it empty.
      text: '',
      fileName: result.fileName,
      isAttachment: true,
      tokenEstimate: result.tokenEstimate,
    },
    kind,
    attachmentId: result.id,
    mimeType: result.mimeType,
    thumbUrl: undefined,
  })
}

/** Upload a file and add it to the context store. Surfaces failures as a
 *  toast so a silent network drop doesn't look like the app swallowed the
 *  file. Returns a promise so callers that want to await (the drop zone)
 *  can, while the paperclip onChange can fire-and-forget. */
export async function handleFileUpload(file: File): Promise<void> {
  try {
    const result = await uploadAttachment(file)
    addAttachmentItem(result)
  } catch (err) {
    console.error('attachment upload failed', err)
    const message =
      err && typeof err === 'object' && 'message' in err
        ? String((err as { message: unknown }).message)
        : 'Could not attach file'
    useUiStore.getState().toast({
      kind: 'error',
      message: `Attachment failed: ${message}`,
    })
  }
}