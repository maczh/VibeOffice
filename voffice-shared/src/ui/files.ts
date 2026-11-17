/* Local file helpers shared by the apps: the browser file picker and a download. */

export function pickFile(accept: string, multiple = false): Promise<File[]> {
  return new Promise((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = accept
    input.multiple = multiple
    input.style.display = 'none'
    document.body.appendChild(input)
    input.addEventListener('change', () => {
      const files = Array.from(input.files || [])
      input.remove()
      resolve(files)
    })
    // a cancelled picker fires no change in some browsers; focus return is the fallback
    window.addEventListener('focus', () => setTimeout(() => { if (input.isConnected) { input.remove(); resolve([]) } }, 500), { once: true })
    input.click()
  })
}

export function downloadBlob(blob: Blob, name: string): void {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = name
  document.body.appendChild(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 10000)
}
