import { ISiteDesign } from "../models/ISiteDesign";
import { IBaseAppState } from "./App";
import { ActionType } from "./IApplicationAction";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ISiteScript } from "../models/ISiteScript";
import { ServiceScope } from "@microsoft/sp-core-library";

export type Page = "Home" | "SiteDesignsList" | "SiteDesignEdition" | "SiteScriptsList" | "SiteScriptEdition";

export interface IApplicationState extends IBaseAppState<ActionType> {
    page: Page;
    currentSiteDesign: ISiteDesign;
    currentSiteScript: ISiteScript;
    allAvailableSiteDesigns: ISiteDesign[];
    allAvailableSiteScripts: ISiteScript[];
    componentContext: WebPartContext;
    serviceScope: ServiceScope;
    isLoading: boolean;
}

export const initialAppState: IApplicationState = {
    page: "Home",
    currentSiteDesign: null,
    currentSiteScript: null,
    componentContext: null,
    serviceScope: null,
    allAvailableSiteDesigns: [],
    allAvailableSiteScripts: [],
    isLoading: false
};