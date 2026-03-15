import { WordProcessorAgent } from "@druide-informatique/antidote-api-js";

import * as utils from "../utils";

export abstract class WordProcessorAgentOnlyOffice extends WordProcessorAgent {
  Asc: IAsc;
  title: string;

  updatingByAntidote: boolean;

  constructor(Asc: IAsc, title: string) {
    super();
    this.Asc = Asc;
    this.title = title;

    this.updatingByAntidote = false;
  }

  abstract updateText(): Promise<void>;

  callCommand<T>(
    func: () => T,
    isClose: boolean = false,
    isCalc: boolean = true,
  ): Promise<T> {
    return utils.callCommand(this.Asc, func, isClose, isCalc);
  }

  executeMethod(
    name: string,
    params: any[]
  ): Promise<any> {
    return utils.executeMethod(this.Asc, name, params);
  }
}
