import {
  ParamsAllowEdit,
  ParamsGetZonesToCorrect,
  ParamsReplace,
  DocumentType,
  TextZoneConnectix,
  WordProcessorConfiguration
} from "@druide-informatique/antidote-api-js";
import { WordProcessorAgentOnlyOffice } from "./base";

export class WordProcessorAgentOnlyOfficeUniversalSelection extends WordProcessorAgentOnlyOffice {
  text: string | null;

  constructor(Asc: IAsc, title: string) {
    super(Asc, title);

    this.text = null;
  }

  configuration(): WordProcessorConfiguration {
    return {
      documentTitle: `${this.title} [selection]`,
      activeMarkup: DocumentType.text,
      carriageReturn: "\r\n"
    };
  }

  applyCorrection(params: ParamsReplace): Promise<void> {
    this.text = (
      this.text!.substring(0, params.positionStartReplace) +
      params.newString +
      this.text!.substring(params.positionReplaceEnd)
    );
    const textArr = this.text!.replace(/(?:\r\n)+$/, "").split(/(?:\r\n){2}|\t/g);
    // console.log("textArr: ", textArr)

    return this.executeMethod("ReplaceTextSmart", [textArr, String.fromCharCode(160)]);
  }


  textZonesAvailable(): boolean {
    if (this.replacingQueue.length > 0
      && !this.mutexQueue.isLocked()
      && !this.mutexDocument.isLocked())
      return false;
    return !!this.text;
  }

  allowEdit(params: ParamsAllowEdit): boolean {
    if (this.text)
      return true;
    return false;
  }

  zonesToCorrect(params: ParamsGetZonesToCorrect): TextZoneConnectix[] {
    return [
      {
        text: this.text ?? "Nothing to correct",
        zoneId: "",
        zoneIsFocused: true
      }
    ]
  }

  updateText(): Promise<void> {
    // console.log("updateText called");
    this.text = null;

    return this.executeMethod("GetSelectedText", [{
      Numbering: false,
      Math: false,
      ParaSeparator: "\r\n\r\n",
      TableRowSeparator: "\r\n\r\n",
      TabSymbol: String.fromCharCode(160)
    }])
    .then((text?: string) => {
      // console.log(`The Text: ${JSON.stringify(text)}`);
      this.text = text ?? null;
    })
  }
}
