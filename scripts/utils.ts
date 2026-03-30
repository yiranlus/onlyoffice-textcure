import Settings from "./settings";

export async function getWebSocketPort(): Promise<number> {
  const portMiminum = 49152;
  const numberOfValidPorts = 12;

  let portWebSocket: number | null = null;

  const testWebSocketPort = (port: number): Promise<boolean> => {
    return new Promise<boolean>((resolve) => {
      let ws = new WebSocket(`ws://localhost:${port}`);
      ws.addEventListener("open", (_event) => {
        portWebSocket = port;
        resolve(true);
      });
      ws.addEventListener("error", (_event) => {
        if (port <= portMiminum + numberOfValidPorts) {
          testWebSocketPort(port + 1).then(resolve);
        } else {
          resolve(false);
        }
      });
    });
  };
  const hasWebSocketPort = await testWebSocketPort(portMiminum);

  if (!hasWebSocketPort || portWebSocket === null)
    throw new Error("No WebSocket port found");

  Settings.setAntidotePort(portWebSocket);
  return portWebSocket;
}

export function applyTranslation(id: string, text: string) {
  const element = document.getElementById(id);
  if (element) {
    element.innerHTML = window.Asc.plugin.tr(text);
  }
}

export function getFullUrl(name: string): string {
  const location = window.location;
  const start = location.pathname.lastIndexOf("/") + 1;
  const file = location.pathname.slice(start);
  return location.href.replace(file, name);
}

export function callCommand<T>(
  func: () => T,
  isClose: boolean = false,
  isCalc: boolean = true,
): Promise<T> {
  return new Promise((resolve) => {
    window.Asc.plugin.callCommand(func, isClose, isCalc, (res: T) => {
      resolve(res);
    });
  });
}

export function executeMethod(name: string, params: any[]): Promise<any> {
  return new Promise((resolve) => {
    window.Asc.plugin.executeMethod(name, params, (res: any) => {
      resolve(res);
    });
  });
}

export async function getDocumentTitle(): Promise<string> {
  switch (window.Asc.plugin.info.editorType) {
    case "word":
      return callCommand(
        () => {
          const oDocument = Api.GetDocument();
          const oDocumentInfo = oDocument.GetDocumentInfo();
          return oDocumentInfo.Title;
        },
        false,
        false,
      );
    case "slide":
      return callCommand(
        () => {
          const oPresentation = Api.GetPresentation();
          const oDocumentInfo = oPresentation.GetDocumentInfo();
          const title = oDocumentInfo.Title;

          return title;
        },
        false,
        false,
      );
    case "cell":
      return callCommand(
        () => {
          const oDocumentInfo = Api.GetDocumentInfo();
          const title = oDocumentInfo.Title;

          return title;
        },
        false,
        false,
      );
  }
  throw new Error("Unsupported editor type");
}
