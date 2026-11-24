/** Office's standard palette, keyed by its stable RGB values. */
export const STANDARD_COLOR_NAMES: Readonly<Record<string, string>> = Object.freeze({
  '#000000': 'Black', '#993300': 'Brown', '#333300': 'Olive Green', '#003300': 'Dark Green',
  '#003366': 'Dark Teal', '#000080': 'Dark Blue', '#333399': 'Indigo', '#333333': 'Gray-80%',
  '#800000': 'Dark Red', '#ff6600': 'Orange', '#808000': 'Dark Yellow', '#008000': 'Green',
  '#008080': 'Teal', '#0000ff': 'Blue', '#666699': 'Blue-Gray', '#808080': 'Gray-50%',
  '#ff0000': 'Red', '#ff9900': 'Light Orange', '#99cc00': 'Lime', '#339966': 'Sea Green',
  '#33cccc': 'Aqua', '#3366ff': 'Light Blue', '#800080': 'Violet', '#969696': 'Gray-40%',
  '#ff00ff': 'Pink', '#ffcc00': 'Gold', '#ffff00': 'Yellow', '#00ff00': 'Bright Green',
  '#00ffff': 'Turquoise', '#00ccff': 'Sky Blue', '#993366': 'Plum', '#c0c0c0': 'Gray-25%',
  '#ff99cc': 'Rose', '#ffcc99': 'Tan', '#ffff99': 'Light Yellow', '#ccffcc': 'Light Green',
  '#ccffff': 'Light Turquoise', '#99ccff': 'Pale Blue', '#cc99ff': 'Lavender', '#ffffff': 'White',
});

export function standardColorName(color: string, translate: (label: string) => string = (label) => label): string {
  const hex = color.toLowerCase();
  const name = STANDARD_COLOR_NAMES[hex];
  return name ? translate(name) : color.toUpperCase();
}
