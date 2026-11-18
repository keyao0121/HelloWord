// The file pickers of the File System Access API, which TypeScript's DOM library leaves out.
// acdert implements them with the system's own dialogs.

interface FilePickerType {
  description?: string
  accept: Record<string, string[]>
}

interface Window {
  showOpenFilePicker?(options?: { types?: FilePickerType[]; multiple?: boolean }): Promise<FileSystemFileHandle[]>
  showSaveFilePicker?(options?: { suggestedName?: string; types?: FilePickerType[] }): Promise<FileSystemFileHandle>
}
