/// <reference types="vite/client" />

import { FUniver } from "@univerjs/facade";
import { Univer } from "@univerjs/core";

declare global  {
  interface Window {
    univer: Univer;
    univerAPI: FUniver;
  }

  // used by demo
  interface HTMLElement{
    tmpListener: any;
    errListener: any;
  }
}
