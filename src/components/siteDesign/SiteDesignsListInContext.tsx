import * as React from "react";
import { useAppContext } from "../../app/App";
import { IApplicationState } from "../../app/ApplicationState";
import { ActionType, IEditSiteDesignActionArgs } from "../../app/IApplicationAction";
import { ISiteDesignsListAllOptionalProps, SiteDesignsList } from "./SiteDesignsList";
import { ISiteDesign } from "../../models/ISiteDesign";


/**
 * This component users the global app context to pass all the site designs to the actual List component
 * @param props 
 */
export const SiteDesignsListInContext = (props: ISiteDesignsListAllOptionalProps) => {
    const [appContext, executeAction] = useAppContext<IApplicationState, ActionType>();

    const onSiteDesignClick = (siteDesign: ISiteDesign) => {
        executeAction("EDIT_SITE_DESIGN", { siteDesign } as IEditSiteDesignActionArgs);
    };

    const onNewSiteDesignAdded = () => {
        const siteDesign: ISiteDesign = {
            Id: null,
            Title: null,
            Description: null,
            Version: 1,
            IsDefault: false,
            PreviewImageAltText: null,
            PreviewImageUrl: null,
            SiteScriptIds: [],
            WebTemplate: ""
        };
        executeAction("EDIT_SITE_DESIGN", { siteDesign } as IEditSiteDesignActionArgs);
    };

    return <SiteDesignsList siteDesigns={appContext.allAvailableSiteDesigns}
        onSiteDesignClicked={onSiteDesignClick}
        onAdd={onNewSiteDesignAdded}  {...props} />;
};