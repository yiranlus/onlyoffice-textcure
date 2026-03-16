import {
  AntidoteConnector,
  ConnectixAgent,
} from "@druide-informatique/antidote-api-js";

import * as utils from "./utils";
import { WordProcessorAgentOnlyOffice } from "./processor-agent/base";
import { WordProcessorAgentOnlyOfficeDocument } from "./processor-agent/document";
import { WordProcessorAgentOnlyOfficeSelection } from "./processor-agent/selection";

((window, undefined) => {
  let wordProcessorAgent: WordProcessorAgentOnlyOffice | null;

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
    if (antidotePort) {
      return Number(antidotePort);
    }

    throw new Error("Antidote port is not set.")
  }

  function launchCorrector() {
    AntidoteConnector.announcePresence();
    const agent = new ConnectixAgent(
      wordProcessorAgent!,
      AntidoteConnector.isDetected() ?
      AntidoteConnector.getWebSocketPort :
      async () => getAntidotePort()
    );

    agent.connectWithAntidote()
      .then(() => agent.launchCorrector())
      .catch(error => {
        const errorDialog = new window.Asc.PluginWindow();
        errorDialog.show(connectionErrorModal);
        window.Asc.plugin.connectionErrorModalId = errorDialog.id;
      })
  }

  window.Asc.plugin.init = () => {
    if (wordProcessorAgent && wordProcessorAgent.isAvailable) {
      // On every selection change
      if (wordProcessorAgent instanceof WordProcessorAgentOnlyOfficeSelection
        && !wordProcessorAgent.updatingByAntidote) {
        setTimeout(() => {
          if (wordProcessorAgent && !wordProcessorAgent.updatingByAntidote) {
            wordProcessorAgent.updateText();
          }
        }, 200);
      }
    } else {
      // Otherwise, create an WordProcessorAgent instance
      utils.callCommand(
        window.Asc,
        () => {
          const oDocument = Api.GetDocument();
          const oDocumentInfo = oDocument.GetDocumentInfo();
          const title = oDocumentInfo.Title;

          const oRange = oDocument.GetRangeBySelect();
          const start = oRange ? oRange.GetStartPos() : null;
          const end = oRange ? oRange.GetEndPos() : null;

          const hasSelection = (start !== end);

          // if (oRange) {
          //   console.log(`oRange Text: "${JSON.stringify(oRange.GetText())}"`);
          //   console.log(range);
          // }

          return { title, hasSelection };
        }
      )
        .then(async ({ title, hasSelection }) => {
          if (hasSelection) {
            wordProcessorAgent = new WordProcessorAgentOnlyOfficeSelection(window.Asc, title);
          } else {
            wordProcessorAgent = new WordProcessorAgentOnlyOfficeDocument(window.Asc, title);
          }
          await wordProcessorAgent.updateText();
        })
        .then(() => {
          launchCorrector();
        });
    }
  };

  window.Asc.plugin.button = (id: string, windowId: string) => {
    if (windowId === window.Asc.plugin.connectionErrorModalId) {
      window.Asc.plugin.executeCommand("close", "");
    }
  };

})(window, undefined);
