export interface MathRun { text: string; kind?: 'normal' | 'operator' | 'number' }
export function ommlToMathText(xml: string): string { return xml.replace(/<m:br\s*\/?>(\s*)/g, '\n').replace(/<[^>]+>/g, '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/\s+/g, ' ').trim() }
export function mathTextToOmml(text: string): string { const esc = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); return `<m:oMathPara><m:oMath><m:r><m:t>${esc}</m:t></m:r></m:oMath></m:oMathPara>` }
