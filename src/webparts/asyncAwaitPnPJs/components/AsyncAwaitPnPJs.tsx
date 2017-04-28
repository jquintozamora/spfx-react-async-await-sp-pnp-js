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


    ////////////////////////////////////////////////////////////////////////
    // enable Logging system
    ////////////////////////////////////////////////////////////////////////
    // we will integrate PnP JS Logging System with SPFx Logging system
    // 1. Logger object = PnP JS Logger
    //    https://github.com/SharePoint/PnP-JS-Core/wiki/Working-With:-Logging
    // 2. Log object = SPFx Logger
    //    https://github.com/SharePoint/sp-dev-docs/wiki/Working-with-the-Logging-API
    ////////////////////////////////////////////////////////////////////////

    // for developement activate Info level
    Logger.activeLogLevel = LogLevel.Info;

    // pnp-js Logger. subscribe a custom listener integrated with SPFx Logging system
    let listener = new FunctionListener((entry: LogEntry) => {
      const componentName: string = (this as any)._reactInternalInstance._currentElement.type.name;

      // instead of using switch we use this easy syntax
      const logLevelConversion = { Verbose: "verbose", Info: "info", Warning: "warn", Error: "error" };

      // we need to trick the message using Error constructor in order to avoit the 24 chara
      const formatedMessage: Error = new Error(`Message: ${entry.message} Data: ${JSON.stringify(entry.data)}`);

      Log[logLevelConversion[LogLevel[entry.level]]](componentName, formatedMessage);
    });
    Logger.subscribe(listener);
  }


  public componentDidMount(): void {
    // this.readItems();
    this.readAllFilesSize("Documents");
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
}
