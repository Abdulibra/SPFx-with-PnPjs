import { WebPartContext } from "@microsoft/sp-webpart-base";
//import pnp and pnp logging system or othre modules if you want.
import { spfi, SPFI, SPFx as spSPFx } from "@pnp/sp";
import { graphfi, GraphFI, SPFx as graphSPFx } from "@pnp/graph";
import { LogLevel, PnPLogging } from "@pnp/logging";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";

let _sp: SPFI = null;
let _graph: GraphFI = null;

export const getSP = (context?: WebPartContext): SPFI => {
  if (context !== null) {
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
    // _sp = spfi().using(spSPFx(context)).using(PnPLogging(LogLevel.Warning));
    _sp = spfi().using(spSPFx(context));
  }
  return _sp;
};

export const getGraph = (context?: WebPartContext): GraphFI => {
  if (context !== null) {
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
    _graph = graphfi()
      .using(graphSPFx(context))
      .using(PnPLogging(LogLevel.Warning));
  }
  return _graph;
};
