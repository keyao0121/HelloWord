// Decoding pictures. In acdert 0.0.9 an <img> loads only while it is in the document, and
// createImageBitmap is not supported, so every picture is decoded through an element in a
// holder that is never shown. A decoded element keeps its pixels after it leaves the tree.

let holder: HTMLDivElement | null = null

/** Decodes a picture. Rejects when the runtime cannot read it. */
export function decodeImage(blob: Blob): Promise<HTMLImageElement> {
  if (!holder) {
    holder = document.createElement('div')
    holder.hidden = true
    document.body.append(holder)
  }
  const url = URL.createObjectURL(blob)
  const image = new Image()
  return new Promise((resolve, reject) => {
    const finish = (): void => {
      URL.revokeObjectURL(url)
      image.remove()
    }
    image.onload = () => {
      finish()
      resolve(image)
    }
    image.onerror = () => {
      finish()
      reject(new Error('无法读取这张图片'))
    }
    holder!.append(image)
    image.src = url
  })
}
