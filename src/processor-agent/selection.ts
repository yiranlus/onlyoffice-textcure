import {
  ParamsAllowEdit,
  ParamsGetZonesToCorrect,
  ParamsReplace,
  TextZoneConnectix,
  WordProcessorConfiguration,
  DocumentType
} from "@druide-informatique/antidote-api-js";
import { WordProcessorAgentOnlyOffice } from "./base";

export type Range = {
  start: number,
  end: number
}

export class WordProcessorAgentOnlyOfficeSelection extends WordProcessorAgentOnlyOffice {
  range: Range;
  text?: string;

  constructor(Asc: IAsc, title: string, range: Range) {
    super(Asc, title);

    this.range = range;
  }

  sessionEnded() {
    this.Asc.plugin.executeCommand("close", "");
    super.sessionEnded();
  }

  configuration(): WordProcessorConfiguration {
    return {
      documentTitle: `${this.title} [selection]`,
      activeMarkup: DocumentType.text
    };
  }

  correctIntoWordProcessor(params: ParamsReplace): boolean {
    this.Asc.scope.selectedRange = this.range;
    this.text = (
      this.text!.substring(0, params.positionStartReplace) +
      params.newString +
      this.text!.substring(params.positionReplaceEnd)
    )

    this.Asc.scope.selectedRange.text = this.text;
    this.callCommand(
      () => {
        const { start, end, text } = Asc.scope.selectedRange;
        const oDocument = Api.GetDocument();
        const oRange = oDocument.GetRange(start, end);

        console.log(`Text to Replace: ${JSON.stringify(text)}`);

        oRange.Select();
        Api.ReplaceTextSmart(text!.split("\r\n\r\n"), String.fromCharCode(160));
      }
    )

    return true;
  }

  allowEdit(params: ParamsAllowEdit): boolean {
    return true;
  }

  textZonesAvailable(): boolean {
    return !!this.text;
  }

  zonesToCorrect(params: ParamsGetZonesToCorrect): TextZoneConnectix[] {
    return [
      {
        text: this.text ? this.text : "Should not be there",
        zoneId: "",
        zoneIsFocused: true
      }
    ]
  }

  updateText(): Promise<void> {
    this.Asc.scope.selectedRange = this.range;

    return this.callCommand(
      () => {
        const oDocument = Api.GetDocument();
        const { start, end } = Asc.scope.selectedRange;

        const oRange = oDocument.GetRange(start, end);
        return oRange.GetText({
          Numbering: false,
          Math: false,
          ParaSeparator: "\r\n\r\n",
          TableCellSeparator: "\r\n",
          TabSymbol: String.fromCharCode(160)
        });
      },
      false,
      false
    )
    .then(text => {
      this.text = text;
      console.log(`Stored text: ${JSON.stringify(this.text)}`);
    });
  }
}
