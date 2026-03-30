import {
  AntidoteConnector,
  ConnectixAgent,
} from "@druide-informatique/antidote-api-js";

import * as utils from "./utils";
import Settings from "./settings";
import { WordProcessorAgentOnlyOffice } from "./processor-agent/base";
import { WordProcessorAgentOnlyOfficeDocument } from "./processor-agent/document";
import { WordProcessorAgentOnlyOfficeUniversalSelection as WordProcessorAgentOnlyOfficeSelection } from "./processor-agent/selection";

export function setupPlugin() {
  let isInitialized = false;
  let wordProcessorAgent: WordProcessorAgentOnlyOffice | null;

  const connectionErrorModal = {
    url: utils.getFullUrl("connection-error.html"), // Same HTML as config variationnt
    description: window.Asc.plugin.tr("Error"),
    isVisual: true,
    EditorsSupport: ["word"],
    isModal: true,
    isInsideMode: false,
    initDataType: "none",
    initData: "",
    size: [350, 150],
    buttons: [
      {
        text: window.Asc.plugin.tr("Close"),
        primary: true,
      },
    ],
  };
  let connectionErrorModalId: string | null;

  const launchCorrector = () => {
    AntidoteConnector.announcePresence();

    if (AntidoteConnector.isDetected()) {
      console.log("Antidote Connector is detected");
    }

    const agent = new ConnectixAgent(
      wordProcessorAgent!,
      Settings.getForceSetPort()
        ? async () => Settings.getAntidotePort()
        : AntidoteConnector.isDetected()
          ? AntidoteConnector.getWebSocketPort
          : utils.getWebSocketPort,
    );

    agent
      .connectWithAntidote()
      .then(() => agent.launchCorrector())
      .catch((error) => {
        const errorDialog = new window.Asc.PluginWindow();
        errorDialog.show(connectionErrorModal);
        connectionErrorModalId = errorDialog.id;

        console.log(error);
      });
  };

  window.Asc.plugin.init = (text: string) => {
    const alternativeText = text.length === 0 ? null : text;

    if (wordProcessorAgent && wordProcessorAgent.isAvailable) {
      // On every selection change

      if (!wordProcessorAgent.updatingByAntidote) {
        if (
          wordProcessorAgent instanceof WordProcessorAgentOnlyOfficeSelection
        ) {
          setTimeout(() => {
            (
              wordProcessorAgent as WordProcessorAgentOnlyOfficeSelection
            ).setAlternativeText(alternativeText);
            if (wordProcessorAgent && !wordProcessorAgent.updatingByAntidote) {
              wordProcessorAgent.updateText();
            }
          }, Settings.getUpdateDelayMS());
        }
      }
    } else {
      // Otherwise, create an WordProcessorAgent instance
      utils
        .getDocumentTitle()
        .then((title) => {
          switch (window.Asc.plugin.info.editorType) {
            case "word":
              if (!alternativeText) {
                wordProcessorAgent = new WordProcessorAgentOnlyOfficeDocument(
                  title,
                );
                break;
              }
            case "slide":
            case "cell":
              wordProcessorAgent = new WordProcessorAgentOnlyOfficeSelection(
                title,
              );
              (
                wordProcessorAgent as WordProcessorAgentOnlyOfficeSelection
              ).setAlternativeText(alternativeText);
              break;
          }
        })
        .then(() => wordProcessorAgent!.updateText())
        .then(launchCorrector);
    }

    isInitialized = true;
  };

  window.Asc.plugin.button = (id: string, windowId: string) => {
    if (connectionErrorModalId && windowId === connectionErrorModalId) {
      window.Asc.plugin.executeCommand("close", "");
    }
  };
}
