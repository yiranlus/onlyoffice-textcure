export function applyTranslation(window: Window, id: string, text: string) {
  const element = document.getElementById(id);
  if (element) {
    element.innerHTML = window.Asc.plugin.tr(text);
  }
}
