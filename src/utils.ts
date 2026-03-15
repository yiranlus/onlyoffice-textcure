export function applyTranslation(Asc: IAsc, id: string, text: string) {
  const element = document.getElementById(id);
  if (element) {
    element.innerHTML = Asc.plugin.tr(text);
  }
}

export function callCommand<T>(
  Asc: IAsc,
  func: () => T,
  isClose: boolean = false,
  isCalc: boolean = true,
): Promise<T> {
  return new Promise(resolve => {
    Asc.plugin.callCommand(func, isClose, isCalc, (res: T) => {
      resolve(res);
    })
  })
}

export function executeMethod(
  Asc: IAsc,
  name: string,
  params: any[]
): Promise<any> {
  return new Promise(resolve => {
    Asc.plugin.executeMethod(name, params, (res: any) => {
      resolve(res);
    })
  })
}
