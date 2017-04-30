import * as React from "react";
import styles from "./AsyncAwaitPnPJs.module.scss";

// create PnP JS response interface for File
interface IResponseFile {
  Length: number;
}

// create PnP JS response interface for Item
interface IResponseItem {
  File: IResponseFile;
  FileLeafRef: string;
  Title: string;
}

// create File item to work with it internally
interface IFile {
  Title: string;
  Name: string;
  Size: number;
}

// import pnp and pnp logging system
import pnp, { Logger, FunctionListener, LogEntry, LogLevel } from "sp-pnp-js";

// import SPFx Logging system
import { Log } from "@microsoft/sp-core-library";

// import React props and state
import { IAsyncAwaitPnPJsProps } from "./IAsyncAwaitPnPJsProps";
import { IAsyncAwaitPnPJsState } from "./IAsyncAwaitPnPJsState";

export default class AsyncAwaitPnPJs extends React.Component<IAsyncAwaitPnPJsProps, IAsyncAwaitPnPJsState> {

  constructor(props: IAsyncAwaitPnPJsProps) {
    super(props);
    // set initial state
    this.state = {
      items: []
    };

    // normally we don't need to bind the functions as we use arrow functions and do automatically the bing
    // https://blog.josequinto.com/2016/12/07/react-use-es6-arrow-functions-in-classes-to-avoid-binding-your-methods-with-the-current-this-object/
    // but using Async function we can't convert it into arrow function, so we do the binding here
    this.readAllFilesSize.bind(this);

    // enable PnP JS Logging integrated with SPFx Logging
    this._enableLogging();
  }

  public componentDidMount(): void {
    this.readAllFilesSize("Documents");
  }

  public render(): React.ReactElement<IAsyncAwaitPnPJsProps> {
    const totalDocs: number = this.state.items.length > 0
      ? this.state.items.reduce<number>((acc: number, item: IFile) => { return (acc + Number(item.Size)); }, 0)
      : 0;
    return (
      <div className={styles.container}>
        <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>
          <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
            <span className="ms-font-xl ms-fontColor-white">Welcome to SharePoint Async Await SP PnP JS Demo!</span>
            <p className="ms-font-l ms-fontColor-white">List of documents:</p>
            <div>
              <div className={styles.row}>
                <div className={styles.left}>Name</div>
                <div className={styles.right}>Size</div>
                <div className={styles.clear + " " + styles.header}></div>
              </div>
              {this.state.items.map((item) => {
                return (
                  <div className={styles.row}>
                    <div className={styles.left}>{item.Name}</div>
                    <div className={styles.right}>{item.Size}</div>
                    <div className={styles.clear}></div>
                  </div>
                );
              })}
              <div className={styles.row}>
                <div className={styles.clear + " " + styles.header}></div>
                <div className={styles.left}>Total: </div>
                <div className={styles.right}>{totalDocs}</div>
                <div className={styles.clear + " " + styles.header}></div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private async readAllFilesSize(libraryName: string): Promise<void> {
    try {
      const response: IResponseItem[] = await pnp.sp
        .web
        .lists
        .getByTitle(libraryName)
        .items
        .select("Title", "FileLeafRef")
        .expand("File/Length")
        .usingCaching()
        .get();
      const items: IFile[] = response.map((item: IResponseItem) => {
        return {
          Title: item.Title,
          Size: item.File.Length,
          Name: item.FileLeafRef
        };
      });
      this.setState({ items });
    } catch (error) {
      // throw new Error(error);
      // do something with State
      this.setState({ items: [] });
    }
  }

  private _enableLogging() {
    ////////////////////////////////////////////////////////////////////////
    // enable Logging system
    ////////////////////////////////////////////////////////////////////////
    // we will integrate PnP JS Logging System with SPFx Logging system
    // 1. Logger object => PnP JS Logger
    //    https://github.com/SharePoint/PnP-JS-Core/wiki/Working-With:-Logging
    // 2. Log object => SPFx Logger
    //    https://github.com/SharePoint/sp-dev-docs/wiki/Working-with-the-Logging-API
    ////////////////////////////////////////////////////////////////////////
    // [PnP JS Logging] activate Info level
    Logger.activeLogLevel = LogLevel.Info;
    // [PnP JS Logging] create a custom FunctionListener to integrate PnP JS and SPFx Logging systems
    let listener = new FunctionListener((entry: LogEntry) => {
      // get React component name
      const componentName: string = (this as any)._reactInternalInstance._currentElement.type.name;
      // mapping betwween PnP JS Log types and SPFx logging methods
      // instead of using switch we use object easy syntax
      const logLevelConversion = { Verbose: "verbose", Info: "info", Warning: "warn", Error: "error" };
      // create Message. Two importante notes here:
      // 1. Use JSON.stringify to output everything. It´s helpful when some internal exception comes thru.
      // 2. Use JavaScript´s Error constructor allows us to output more than 100 characters using SPFx logging
      const formatedMessage: Error = new Error(`Message: ${entry.message} Data: ${JSON.stringify(entry.data)}`);
      // [SPFx Logging] Calculate method to invoke verbose, info, warn or error
      const method = logLevelConversion[LogLevel[entry.level]];
      // [SPFx Logging] Call SPFx Logging system with the message received from PnP JS Logging
      Log[method](componentName, formatedMessage);
    });
    // [PnP JS Logging] Once create the custom listerner we should subscribe to it
    Logger.subscribe(listener);
  }
}
