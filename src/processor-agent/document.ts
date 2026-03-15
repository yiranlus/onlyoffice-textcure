import {
  ParamsReplace,
  ParamsAllowEdit,
  WordProcessorConfiguration,
  ParamsGetZonesToCorrect,
  TextZoneConnectix,
  DocumentType,
} from "@druide-informatique/antidote-api-js";
import { Mutex } from "async-mutex";

import { WordProcessorAgentOnlyOffice as WordProcessorAgentOnlyOfficeBase } from "./base";

type Paragraph = {
  globalPos: number,
  text?: string
}

export class EmptyDataError extends Error {
  constructor() {
    super("Data is empty");
    this.name = 'EmptyDataError';
    Object.setPrototypeOf(this, EmptyDataError.prototype);
  }
}

export class WordProcessorAgentOnlyOfficeDocument extends WordProcessorAgentOnlyOfficeBase {
  paragraphs: Paragraph[] | null;

  replacingQueue: ParamsReplace[];
  mutexQueue: Mutex;
  mutexDocument: Mutex;

  mutexUpdateText: Mutex;

  constructor(Asc: any, title: string) {
    super(Asc, title);

    this.paragraphs = null;

    this.replacingQueue = [];
    this.mutexQueue = new Mutex();
    this.mutexDocument = new Mutex();
    this.mutexUpdateText = new Mutex();
  }

  sessionEnded() {
    this.Asc.plugin.executeCommand("close", "");
    super.sessionEnded();
  }

  findIndex(pos: number, eager: boolean = false): number {
    if (!this.paragraphs) {
      throw new Error("Data is empty");
    }

    let elementIndex = 0;
    if (eager) {
      while (
        elementIndex + 1 < this.paragraphs.length &&
        this.paragraphs[elementIndex + 1].globalPos <= pos)
        elementIndex++;
    } else {
      while (
        elementIndex + 1 < this.paragraphs.length &&
        this.paragraphs[elementIndex + 1].globalPos < pos)
        elementIndex++;
    }

    return elementIndex;
  }

  applyCorrection(params: ParamsReplace): Promise<void> {
    // Waiting to previous action to finish
    // console.log("ParasReplace: ", params);

    let elementIndex = this.findIndex(params.positionStartReplace, true);
    const globalPos = this.paragraphs![elementIndex].globalPos;

    let text = this.paragraphs![elementIndex].text!;
    let newText = (
      text.substring(0, params.positionStartReplace - globalPos) +
      params.newString +
      text.substring(params.positionReplaceEnd - globalPos)
    ).replace(/(\r\n)*$/, "");

    // console.log(`${elementIndex} => "${newText}"`);

    this.Asc.scope.paramsReplace = { elementIndex, text: newText };

    return this.callCommand(
      () => {
        const { elementIndex, text } = Asc.scope.paramsReplace;

        var oDocument = Api.GetDocument();
        var oElement = oDocument.GetElement(elementIndex);

        var oldText = oElement.GetText({ "Numbering": false }).replace(/(\r\n)*$/, "");

        oElement.Select();
        Api.ReplaceTextSmart([text]);

        const newText = oElement.GetText({ "Numbering": false }).replace(/(\r\n)*$/, "");

        return {
          text: newText,
          diff: newText.length - oldText.length
        }
      },
      false,
      true
    )
    .then(res => {
      this.paragraphs![elementIndex].text = res.text;
      for (let i = elementIndex + 1; i < this.paragraphs!.length; i++) {
        this.paragraphs![i].globalPos += res.diff;
      }
    })
    .catch(err => console.log(err));
  }

  configuration(): WordProcessorConfiguration {
    return {
      documentTitle: this.title,
      activeMarkup: DocumentType.text
    };
  }

  allowEdit(params: ParamsAllowEdit): boolean {
    return true;
  }

  textZonesAvailable(): boolean {
    if (this.replacingQueue.length > 0
      && !this.mutexQueue.isLocked()
      && !this.mutexDocument.isLocked())
      return false;
    return !!this.paragraphs;
  }

  zonesToCorrect(_params: ParamsGetZonesToCorrect): TextZoneConnectix[] {
    const text = this.paragraphs!.map(el => el.text).join("\r\n\r\n");
    return [{
      text,
      zoneId: "",
      zoneIsFocused: true,
    }];
  }

  updateText(): Promise<void> {
    console.log("_updateParagraphs called");
    this.paragraphs = null;

    return this.callCommand(
      () => {
        const oDocument = Api.GetDocument();

        let paragraphs: Paragraph[] = [], globalPos = 0;
        for (let i = 0; i < oDocument.GetElementsCount(); i++) {
          const element = oDocument.GetElement(i);
          if (element.GetClassType() === "paragraph") {
            const text = element.GetText({ "Numbering": false }).replace(/(\r\n)*$/, "");
            paragraphs.push({ globalPos, text });
            globalPos += text.length;
          } else {
            paragraphs.push({ globalPos });
          }
        }

        return paragraphs;
      },
      false,
      false
    )
    .then(paragraphs => {
      for (let i = 1; i < paragraphs.length; i++) {
        paragraphs[i].globalPos += 4 * i;
      }
      this.paragraphs = paragraphs;
    })
    .catch(err => console.log(err));
  }
}
