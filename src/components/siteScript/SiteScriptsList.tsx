import * as React from "react";
import {
    DocumentCard,
    DocumentCardPreview,
    DocumentCardDetails,
    DocumentCardTitle,
    IDocumentCardPreviewProps,
    DocumentCardType
} from 'office-ui-fabric-react/lib/DocumentCard';
import { ImageFit } from 'office-ui-fabric-react/lib/Image';
import { ISize } from 'office-ui-fabric-react/lib/Utilities';
import { GridLayout } from "@pnp/spfx-controls-react/lib/GridLayout";
import { ActionType, IEditSiteScriptActionArgs } from "../../app/IApplicationAction";
import styles from "./SiteScriptsList.module.scss";
import { useAppContext } from "../../app/App";
import { IApplicationState } from "../../app/ApplicationState";
import { ISiteScript } from "../../models/ISiteScript";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { Icon } from "office-ui-fabric-react/lib/Icon";
import { SiteScriptSchemaServiceKey } from "../../services/siteScriptSchema/SiteScriptSchemaService";
import { Link } from "office-ui-fabric-react/lib/Link";
import { useState } from "react";
import { Panel } from "office-ui-fabric-react/lib/Panel";
import { SiteDesignsServiceKey } from "../../services/siteDesigns/SiteDesignsService";

export interface ISiteScriptsListProps {
    preview?: boolean;
    addNewDisabled?: boolean;
}

const PREVIEW_ITEMS_COUNT = 3;

export const SiteScriptsList = (props: ISiteScriptsListProps) => {

    const [appContext, executeAction] = useAppContext<IApplicationState, ActionType>();
    // Get services instances
    const siteScriptSchemaService = appContext.serviceScope.consume(SiteScriptSchemaServiceKey);
    const siteDesignsService = appContext.serviceScope.consume(SiteDesignsServiceKey);

    const [isAdding, setIsAdding] = useState<boolean>(false);

    const onSiteScriptClick = (siteScript: ISiteScript) => {
        executeAction("EDIT_SITE_SCRIPT", { siteScript } as IEditSiteScriptActionArgs);
    };

    const onAddNewScript = () => {
        setIsAdding(true);
    };

    const onAddNewBlankScript = () => {
        const newSiteScriptContent = siteScriptSchemaService.getNewSiteScript();
        const siteScript: ISiteScript = {
            Id: null,
            Title: null,
            Description: null,
            Version: 1,
            Content: newSiteScriptContent
        };
        executeAction("EDIT_SITE_SCRIPT", { siteScript } as IEditSiteScriptActionArgs);
        setIsAdding(false);
    };

    const onAddNewScriptFromSite = (siteUrl: string) => {
        siteDesignsService.getSiteScriptFromWeb(siteUrl, {
            includeBranding: true, includeLinksToExportedItems: true,
            includeRegionalSettings: true,
            includeSiteExternalSharingCapability: true,
            includeTheme: true,
            includeLists: ['teamworkmetadata']
        }).then(result => {
            const siteScript: ISiteScript = {
                Id: null,
                Title: null,
                Description: null,
                Version: 1,
                Content: result.JSON
            };
            executeAction("EDIT_SITE_SCRIPT", { siteScript } as IEditSiteScriptActionArgs);
            setIsAdding(false);
        });
    };

    const onAddNewScriptFromList = (listUrl: string) => {
        siteDesignsService.getSiteScriptFromList(listUrl).then(siteScriptContent => {
            const siteScript: ISiteScript = {
                Id: null,
                Title: null,
                Description: null,
                Version: 1,
                Content: siteScriptContent
            };
            executeAction("EDIT_SITE_SCRIPT", { siteScript } as IEditSiteScriptActionArgs);
            setIsAdding(false);
        });
    };

    const renderSiteScriptGridItem = (siteScript: ISiteScript, finalSize: ISize, isCompact: boolean): JSX.Element => {
        if (!siteScript) {
            // If site script is not set, it is the Add new tile
            return <div
                className={styles.add}
                data-is-focusable={true}
                role="listitem"
                aria-label={"Add a new Site Script"}
            >
                <DocumentCard
                    type={isCompact ? DocumentCardType.compact : DocumentCardType.normal}
                    onClick={(ev: React.SyntheticEvent<HTMLElement>) => onAddNewScript()}>
                    <div className={styles.iconBox}>
                        <div className={styles.icon}>
                            <Icon iconName="Add" />
                        </div>
                    </div>
                </DocumentCard>
            </div>;
        }

        return <div
            data-is-focusable={true}
            role="listitem"
            aria-label={siteScript.Title}
        >
            <DocumentCard
                type={isCompact ? DocumentCardType.compact : DocumentCardType.normal}
                onClick={(ev: React.SyntheticEvent<HTMLElement>) => onSiteScriptClick(siteScript)}>
                <div className={styles.iconBox}>
                    <div className={styles.icon}>
                        <Icon iconName="Script" />
                    </div>
                </div>
                <DocumentCardDetails>
                    <DocumentCardTitle
                        title={siteScript.Title}
                        shouldTruncate={true}
                    />
                </DocumentCardDetails>
            </DocumentCard>
        </div>;
    };

    const renderAddNewGridItem = (createScript: { from: "BLANK" | "WEB" | "LIST" }, finalSize: ISize, isCompact: boolean): JSX.Element => {

        // TODO Review the icons
        const iconName = createScript.from == "BLANK" ? "Add" : "Script";
        const title = createScript.from == "BLANK"
            ? "Add a blank Site Script"
            : createScript.from == "WEB"
                ? "Add a Site Script from an existing site"
                : "Add a Site Script from an existing list";
        return <div
            data-is-focusable={true}
            role="listitem"
            aria-label={`Add (${createScript.from})`}
        >
            <DocumentCard
                type={isCompact ? DocumentCardType.compact : DocumentCardType.normal}
                onClick={(ev: React.SyntheticEvent<HTMLElement>) => {
                    if (createScript.from == "BLANK") {
                        onAddNewBlankScript();
                    } else if (createScript.from == "WEB") {
                        onAddNewScriptFromSite("https://pvxdev.sharepoint.com/sites/teamworkadmin");
                    } else if (createScript.from == "LIST") {
                        onAddNewScriptFromList("https://pvxdev.sharepoint.com/sites/teamworkadmin/teamworkmetadata");
                    }
                }}>
                <div className={styles.iconBox}>
                    <div className={styles.icon}>
                        <Icon iconName={iconName} />
                    </div>
                </div>
                <DocumentCardDetails>
                    <DocumentCardTitle
                        title={title}
                        shouldTruncate={true}
                    />
                </DocumentCardDetails>
            </DocumentCard>
        </div>;
    };

    let items = [...appContext.allAvailableSiteScripts];
    if (props.preview) {
        items = items.slice(0, PREVIEW_ITEMS_COUNT);
    }
    if (!props.addNewDisabled) {
        items.push(null);
    }
    const seeMore = props.preview && appContext.allAvailableSiteScripts.length > PREVIEW_ITEMS_COUNT;
    return <div className={styles.SiteDesignsList}>
        <Panel isOpen={isAdding} title="Add a new Site Script">
            <GridLayout
                ariaLabel="Add a new script"
                items={[{ from: "BLANK" }, { from: "WEB" }, { from: "LIST" }]}
                onRenderGridItem={renderAddNewGridItem}
            />
        </Panel>
        <div className={styles.row}>
            <div className={styles.column}>
                <GridLayout
                    ariaLabel="List of Site Scripts."
                    items={[...appContext.allAvailableSiteScripts, null]}
                    onRenderGridItem={renderSiteScriptGridItem}
                />
                {seeMore && <div className={styles.seeMore}>
                    {`There are more than ${PREVIEW_ITEMS_COUNT} available Site Scripts on your tenant. `}
                    <Link onClick={() => executeAction("GO_TO", { page: "SiteScriptsList" })}>See all Site Scripts</Link>
                </div>}
            </div>
        </div>
    </div>;
};