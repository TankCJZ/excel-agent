export function attachmentContentDisposition(fileName: string) {

  const normalized = fileName.replace(/[\r\n"\\/]/g, '-').trim() || 'excelgen-result.xlsx'
  const extension = normalized.match(/\.[a-zA-Z0-9]{1,10}$/)?.[0] || '.xlsx'
  const asciiFallback = `excelgen-result${extension.toLowerCase()}`
  const encoded = encodeURIComponent(normalized).replace(/[!'()*]/g, character => (
    `%${character.charCodeAt(0).toString(16).toUpperCase()}`
  ))
  return `attachment; filename="${asciiFallback}"; filename*=UTF-8''${encoded}`
}
