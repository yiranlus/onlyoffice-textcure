import { applyTranslation } from "./utils";

((window, undefined) => {
  let inputAntidotePort: HTMLInputElement | null;
  let inputForceSetPort: HTMLInputElement | null;
  let inputUpdateDelayMS: HTMLInputElement | null;

  window.Asc.plugin.init = function () {
    inputAntidotePort = document.getElementById("antidotePort") as HTMLInputElement;
    inputUpdateDelayMS = document.getElementById("updateDelayMS") as HTMLInputElement;
    inputForceSetPort = document.getElementById("forceSetPort") as HTMLInputElement;

    if (inputAntidotePort) {
      const antidotePort = window.localStorage.getItem("ANTIDOTE_PORT");
      inputAntidotePort.value = antidotePort ?? "49152";
    }
    if (inputUpdateDelayMS) {
      const updateDelayMS = window.localStorage.getItem("UPDATE_DELAY_MS");
      inputUpdateDelayMS.value = updateDelayMS ?? "200";
    }
    if (inputForceSetPort) {
      const forceSetPort = window.localStorage.getItem("FORCE_SET_PORT");
      if (forceSetPort === "true")
        inputForceSetPort.checked = true;
    }
  };

  window.Asc.plugin.button = (id: string, windowId: string) => {
    const antidotePort = Number(inputAntidotePort?.value);
    const updateDelayMS = Number(inputUpdateDelayMS?.value);
    const forceSetPort = inputForceSetPort?.checked;

    // Send value back to main plugin context (optional)
    window.localStorage.setItem("ANTIDOTE_PORT", antidotePort.toString());
    window.localStorage.setItem("UPDATE_DELAY_MS", updateDelayMS.toString());
    window.localStorage.setItem("FORCE_SET_PORT", forceSetPort?"true":"false");

    window.Asc.plugin.executeCommand("close", "");
  };

  window.Asc.plugin.onTranslate = () => {
    applyTranslation(window.Asc, "lblAntidotePort", "Websocket Port:");
    applyTranslation(window.Asc, "lblUpdateDelayMS", "Update Delay (ms):");
    applyTranslation(window.Asc, "lblForceSetPort", "Ignore Antidote Connector");
  }
})(window, undefined);
