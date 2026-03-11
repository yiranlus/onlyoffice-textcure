import {
  AntidoteConnector,
  ConnectixAgent,
} from "@druide-informatique/antidote-api-js";
import { WordProcessorAgentOnlyOfficeDocument } from "./processor-agent";

(function(window, undefined){
  const wordProcessorAgent = new WordProcessorAgentOnlyOfficeDocument(window.Asc);

  function getFullUrl(name: string): string {
    const location = window.location;
    const start = location.pathname.lastIndexOf("/") + 1;
    const file = location.pathname.slice(start);
    return location.href.replace(file, name);
  }

  const connectionErrorModal = {
    url: getFullUrl("connection-error.html"),  // Same HTML as config variationnt
    description: window.Asc.plugin.tr("Error"),
    isVisual: true,
    EditorsSupport: ["word"],
    isModal : true,
    isInsideMode : false,
    initDataType : "none",
    initData : "",
    size: [350, 150],
    buttons: [
      {
        text: window.Asc.plugin.tr("Close"),
        primary: true
      }
    ]
  };

  function getAntidotePort() {
    const antidotePort = localStorage.getItem("ANTIDOTE_PORT");
    // console.log("antidotePort: ", antidotePort)
    if (antidotePort) {
      return Number(antidotePort);
    }

    throw new Error("Antidote port is not set.")
  }

  function launchCorrector() {
    AntidoteConnector.announcePresence();
    console.log("Status of AntidoteConnector: ", AntidoteConnector.isDetected());

    const agent = new ConnectixAgent(
      wordProcessorAgent,
      AntidoteConnector.isDetected() ?
      AntidoteConnector.getWebSocketPort :
      async () => getAntidotePort()
    );

    agent.connectWithAntidote()
      .then(() => agent.launchCorrector())
      .catch(error => {
        window.Asc.plugin.executeMethod("ShowWindow", ["iframe_asc.{E649827B-6CD5-477F-A7A7-C6952C813ADE}", connectionErrorModal]);

        console.log("Error Encountered: ", error)
      })
  }

  window.Asc.plugin.init = () => {
    window.Asc.plugin.attachEditorEvent("onDocumentContentReady", () => {
      wordProcessorAgent.updateParagraphs();
    });

    window.Asc.plugin.attachEditorEvent("onParagraphText", (data: any) => {
      if (!wordProcessorAgent.updateByAntidote) {
        wordProcessorAgent.updateParagraphs();
      }
    });

    launchCorrector();
  };

  window.Asc.plugin.button = function(id: string, windowId: string) {
    if (windowId === "iframe_asc.{E649827B-6CD5-477F-A7A7-C6952C813ADE}") {
      window.Asc.plugin.executeCommand("close", "");
    }
  };

})(window, undefined);
