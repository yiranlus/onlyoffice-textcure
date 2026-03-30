import { applyTranslation } from "./utils";

let antidotePort: number | undefined;
let updateDelayMS: number | undefined;
let forceSetPort: boolean | undefined;

function loadSettings() {
  const _antidotePort = window.localStorage.getItem("ANTIDOTE_PORT");
  antidotePort = _antidotePort ? Number(_antidotePort) : 49152;

  const _updateDelayMS = window.localStorage.getItem("UPDATE_DELAY_MS");
  updateDelayMS = _updateDelayMS ? Number(_updateDelayMS) : 200;

  const _forceSetPort = window.localStorage.getItem("FORCE_SET_PORT");
  forceSetPort = _forceSetPort === "true" ? true : false;
}

function getAntidotePort(): number {
  if (antidotePort) return antidotePort;

  loadSettings();
  return antidotePort!;
}

function setAntidotePort(port: number) {
  antidotePort = port;
  window.localStorage.setItem("ANTIDOTE_PORT", port.toString());
}

function getUpdateDelayMS(): number {
  if (updateDelayMS) return updateDelayMS;

  loadSettings();
  return updateDelayMS!;
}

function setUpdateDelayMS(delay: number) {
  updateDelayMS = delay;
  window.localStorage.setItem("UPDATE_DELAY_MS", delay.toString());
}

function getForceSetPort(): boolean {
  if (forceSetPort) return forceSetPort;

  loadSettings();
  return forceSetPort!;
}

function setForceSetPort(force: boolean) {
  forceSetPort = force;
  window.localStorage.setItem("FORCE_SET_PORT", force.toString());
}

const Settings = {
  getAntidotePort,
  setAntidotePort,
  getUpdateDelayMS,
  setUpdateDelayMS,
  getForceSetPort,
  setForceSetPort,
};
export default Settings;

export function setupPlugin() {
  let inputAntidotePort: HTMLInputElement | null;
  let inputForceSetPort: HTMLInputElement | null;
  let inputUpdateDelayMS: HTMLInputElement | null;

  window.Asc.plugin.init = function () {
    inputAntidotePort = document.getElementById(
      "antidotePort",
    ) as HTMLInputElement;
    inputUpdateDelayMS = document.getElementById(
      "updateDelayMS",
    ) as HTMLInputElement;
    inputForceSetPort = document.getElementById(
      "forceSetPort",
    ) as HTMLInputElement;

    if (inputAntidotePort) {
      inputAntidotePort.value = Settings.getAntidotePort().toString();
    }
    if (inputUpdateDelayMS) {
      inputUpdateDelayMS.value = Settings.getUpdateDelayMS().toString();
    }
    if (inputForceSetPort) {
      inputForceSetPort.checked = Settings.getForceSetPort();
    }
  };

  window.Asc.plugin.button = (_id: string, _windowId: string) => {
    const _antidotePort = Number(inputAntidotePort!.value);
    const _updateDelayMS = Number(inputUpdateDelayMS!.value);
    const _forceSetPort = inputForceSetPort!.checked;

    // Send value back to main plugin context (optional)
    Settings.setAntidotePort(_antidotePort);
    Settings.setUpdateDelayMS(_updateDelayMS);
    Settings.setForceSetPort(_forceSetPort);

    window.Asc.plugin.executeCommand("close", "");
  };

  window.Asc.plugin.onTranslate = () => {
    applyTranslation("lblAntidotePort", "Websocket Port:");
    applyTranslation("lblUpdateDelayMS", "Update Delay (ms):");
    applyTranslation("lblForceSetPort", "Disable Automatic Port Detection");
  };
}
