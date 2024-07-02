import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/sp/presets/all";
import "@pnp/sp/profiles";
import { graphfi, GraphFI, SPFx as graphSPFx } from '@pnp/graph';
import "@pnp/graph/groups";
import "@pnp/graph/users";
import { objectDefinedNotNull, stringIsNullOrEmpty } from '@pnp/core/util';

var _sp: SPFI | null = null;

export const getSP = (context?: WebPartContext): SPFI => {
  if (!!context) {
    _sp = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
  }
  return _sp!;
};

export const getGraph = (context: WebPartContext, siteUrl?: string): GraphFI => {
  let _graph: GraphFI | null = null;

  if (_graph === null && context !== null && objectDefinedNotNull(context.pageContext)) {
    _graph = (stringIsNullOrEmpty(siteUrl))
        ? graphfi().using(graphSPFx(context)).using(PnPLogging(LogLevel.Warning))
        : graphfi(siteUrl).using(graphSPFx(context)).using(PnPLogging(LogLevel.Warning))
}
  return _graph!;
};