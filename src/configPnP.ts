import { WebPartContext } from "@microsoft/sp-webpart-base";
import { LogLevel, PnPLogging } from "@pnp/logging";
import { spfi, SPFI, SPFx } from "@pnp/sp";

// import sp list
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";

// Create a configuration file of PnP
// Put all the configuration that's going to be reused throughout the whole project
//  Which basically says here's the context that I'm using this obj to access all the files in the list
let _sp: SPFI = null;

// Using sp through out the web part
// getSP is used to get all the SharePoint context, the parameter is WebPartContext
export const getSP = (context?: WebPartContext): SPFI => {
    if (_sp === null && context !== null) {
        // You must add the @pnp/logging package to include the PnPLogging
        // The LogLevel set's at what level a message will be written to the log
        _sp = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
    }
    return _sp;
};
