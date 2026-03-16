import {
  ParamsAllowEdit,
  ParamsGetZonesToCorrect,
  ParamsReplace,
  TextZoneConnectix,
  WordProcessorConfiguration,
  DocumentType
} from "@druide-informatique/antidote-api-js";
import { WordProcessorAgentOnlyOffice } from "./base";
import { Mutex } from "async-mutex";

export type Range = {
  start: number,
  end: number
}

export class WordProcessorAgentOnlyOfficeSelection extends WordProcessorAgentOnlyOffice {
  range: Range | null;
  text: string | null;

  constructor(Asc: IAsc, title: string) {
    super(Asc, title);

    this.range = null;
    this.text = null;
  }

  configuration(): WordProcessorConfiguration {
    return {
      documentTitle: `${this.title} [selection]`,
      activeMarkup: DocumentType.text,
      carriageReturn: "\r\n"
    };
  }

  applyCorrection(params: ParamsReplace) {
    this.text = (
      this.text!.substring(0, params.positionStartReplace) +
      params.newString +
      this.text!.substring(params.positionReplaceEnd)
    )
    this.Asc.scope.selectedRange = { ...this.range!, text: this.text };

    return this.callCommand(
      () => {
        const { start, end, text } = Asc.scope.selectedRange;
        const oDocument = Api.GetDocument();
        const oRange = oDocument.GetRange(start, end);

        const textArr = text!.replace(/(?:\r\n)+$/, "").split(/(?:\r\n)+|\t/g);

        // console.log(`Text to Replace: ${JSON.stringify(text)}`);
        // console.log(`Text Array to Replace: ${JSON.stringify(textArr)}`);

        oRange.Select();
        Api.ReplaceTextSmart(textArr, String.fromCharCode(160));

        // console.log("Updated Range.");
        // console.log("start: ", oRange.GetStartPos(), "end: ", oRange.GetEndPos())

        return {
          start: oRange.GetStartPos(),
          end: oRange.GetEndPos(),
        } as Range;
      },
      false,
      true
    )
    .then(range => {
      this.range = range;
      // console.log("Updated range: ", this.range);
    });
  }

  allowEdit(params: ParamsAllowEdit): boolean {
    if (this.range)
      return true;
    return false;
  }

  textZonesAvailable(): boolean {
    if (this.replacingQueue.length > 0
      && !this.mutexQueue.isLocked()
      && !this.mutexDocument.isLocked())
      return false;
    return !!this.text;
  }

  zonesToCorrect(_params: ParamsGetZonesToCorrect): TextZoneConnectix[] {
    return [
      {
        text: this.range? this.text!: "Please select some text",
        zoneId: "",
        zoneIsFocused: true
      }
    ]
  }

  updateText(): Promise<void> {
    // console.log("updateText called");
    this.text = null;

    return this.callCommand(
      () => {
        const oDocument = Api.GetDocument();

        const oRange = oDocument.GetRangeBySelect();
        const start = oRange ? oRange.GetStartPos() : null;
        const end = oRange ? oRange.GetEndPos() : null;

        const range = (start === end) ? null : { start, end };

        const text = oRange.GetText({
          Numbering: false,
          Math: false,
          ParaSeparator: "\r\n\r\n",
          TableRowSeparator: "\r\n\r\n",
          TabSymbol: String.fromCharCode(160)
        });

        return { range, text };
      },
      false,
      false
    )
    .then(({range, text}) => {
      this.range = range as (Range | null);
      this.text = text;
      // console.log(`Stored text: ${JSON.stringify(this.text)}`);
    });
  }
}
