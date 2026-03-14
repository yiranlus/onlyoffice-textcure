interface IAsc {
  PluginWindow: new () => any;

  plugin: any;
  scope: {
    paramsReplace: {
      elementIndex: number,
      text: string
    }
  }
}

declare global {
  /**
  * Document Builder API - available inside Asc.plugin.callCommand()
  * Use (Api as any) in code
  */
  var Asc: IAsc;
  var Api: any;
  interface Window {
    /**
    * OnlyOffice plugin globals - intentionally untyped for simplicity
    * Use (window.Asc as any) in code
    */
    Asc: IAsc;
  }
}

export {};
