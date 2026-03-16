import { ParamsReplace, WordProcessorAgent } from "@druide-informatique/antidote-api-js";

import * as utils from "../utils";
import { E_CANCELED, Mutex } from "async-mutex";

export abstract class WordProcessorAgentOnlyOffice extends WordProcessorAgent {
  Asc: IAsc;
  title: string;

  replacingQueue: ParamsReplace[];
  updatingByAntidote: boolean;

  mutexQueue: Mutex;
  mutexDocument: Mutex;
  mutexUpdateText: Mutex;

  isAvailable: boolean;

  constructor(Asc: IAsc, title: string) {
    super();
    this.Asc = Asc;
    this.title = title;

    this.replacingQueue = [];
    this.mutexQueue = new Mutex();

    this.updatingByAntidote = false;
    this.mutexDocument = new Mutex();
    this.mutexUpdateText = new Mutex();

    this.isAvailable = true;
  }

  sessionStarted(): void {
    this.Asc.plugin.attachEditorEvent("onParagraphText", (data: any) => {
      if (!this.updatingByAntidote) {

        // Check if currently the text is updated by Antidote,
        // if not, wait sometime and then recheck to ensure that the
        // replacingQueue is empty
        setTimeout(() => {
          if (!this.updatingByAntidote) {
            // console.log("From onParagraphText", data)
            this!._internalUpdateText();
          }
        }, 200);
      }
    });
  }

  sessionEnded(): void {
    super.sessionEnded();
    this.Asc.plugin.detachEditorEvent("onParagraphText");
    this.Asc.plugin.executeCommand("close", "");

    this.isAvailable = false;
  }

  _internalUpdateText() {
    // Only the last call to updateText is effective, so here canceling all the
    // previous call that hasn't aquire the lock.
    this.mutexUpdateText.cancel();

    return this.mutexUpdateText.runExclusive(
      () => this.mutexDocument.runExclusive(
        () => this.updateText()
      )
    )
    .catch(err => {
      if (err === E_CANCELED) { return; }
      throw err;
    });
  }

  abstract updateText(): Promise<void>;


  correctIntoWordProcessor(params: ParamsReplace): boolean {
    this.mutexQueue.runExclusive(() => {
      // console.log("Locking Queue");
      this.replacingQueue.push(params);
    })
      .then(() => {
        // console.log("Calling Apply Corrections")
        this.applyCorrections();
      });

    return true;
  }

  async applyCorrections() {
    await this.mutexDocument.runExclusive(async () => {
      this.updatingByAntidote = true;

      let params = await this.mutexQueue.runExclusive(() => {
        // console.log("Retriving Item from the Queue")
        return this.replacingQueue.shift();
      });

      while (params) {
        // console.log("Applying a Correction")
        await this.applyCorrection(params!);

        params = await this.mutexQueue.runExclusive(() => {
          // console.log("Retriving Item from the Queue")
          return this.replacingQueue.shift();
        })
      }

      this.updatingByAntidote = false;
    });
  }

  abstract applyCorrection(params: ParamsReplace): Promise<void>;


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
