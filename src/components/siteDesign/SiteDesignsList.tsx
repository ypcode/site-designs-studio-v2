import * as React from "react";
import {
    DocumentCard,
    DocumentCardActivity,
    DocumentCardPreview,
    DocumentCardDetails,
    DocumentCardTitle,
    IDocumentCardPreviewProps,
    DocumentCardLocation,
    DocumentCardType
} from 'office-ui-fabric-react/lib/DocumentCard';
import { ImageFit } from 'office-ui-fabric-react/lib/Image';
import { ISize } from 'office-ui-fabric-react/lib/Utilities';
import { GridLayout } from "@pnp/spfx-controls-react/lib/GridLayout";
import { IEditSiteDesignActionArgs, IGoToActionArgs, ActionType } from "../../app/IApplicationAction";
import { ISiteDesign } from "../../models/ISiteDesign";
import styles from "./SiteDesignsList.module.scss";
import { useAppContext } from "../../app/App";
import { IApplicationState } from "../../app/ApplicationState";
import { Icon } from "office-ui-fabric-react/lib/Icon";
import { Link } from "office-ui-fabric-react/lib/Link";

export interface ISiteDesignsListProps {
    preview?: boolean;
    addNewDisabled?: boolean;
}

const PREVIEW_ITEMS_COUNT = 3;

export const SiteDesignsList = (props: ISiteDesignsListProps) => {

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

    const renderGridItem = (siteDesign: ISiteDesign, finalSize: ISize, isCompact: boolean): JSX.Element => {

        if (!siteDesign) {
            // If site script is not set, it is the Add new tile
            return <div
                className={styles.add}
                data-is-focusable={true}
                role="listitem"
                aria-label={"Add a new Site Design"}
            >
                <DocumentCard
                    type={isCompact ? DocumentCardType.compact : DocumentCardType.normal}
                    onClick={(ev: React.SyntheticEvent<HTMLElement>) => onNewSiteDesignAdded()}>
                    <div className={styles.iconBox}>
                        <div className={styles.icon}>
                            <Icon iconName="Add" />
                        </div>
                    </div>
                </DocumentCard>
            </div>;
        }


        const previewProps: IDocumentCardPreviewProps = {
            previewImages: [
                {
                    previewImageSrc: siteDesign.PreviewImageUrl,
                    imageFit: ImageFit.cover,
                    height: 130
                }
            ]
        };

        return <div
            data-is-focusable={true}
            role="listitem"
            aria-label={siteDesign.Title}
        >
            <DocumentCard
                type={isCompact ? DocumentCardType.compact : DocumentCardType.normal}
                onClick={(ev: React.SyntheticEvent<HTMLElement>) => onSiteDesignClick(siteDesign)}>
                <DocumentCardPreview {...previewProps} />
                <DocumentCardDetails>
                    <DocumentCardTitle
                        title={siteDesign.Title}
                        shouldTruncate={true}
                    />
                </DocumentCardDetails>
            </DocumentCard>
        </div>;
    };

    let items = [...appContext.allAvailableSiteDesigns];
    if (props.preview) {
        items = items.slice(0, PREVIEW_ITEMS_COUNT);
    }
    if (!props.addNewDisabled) {
        items.push(null);
    }
    const seeMore = props.preview && appContext.allAvailableSiteDesigns.length > PREVIEW_ITEMS_COUNT;
    return <div className={styles.SiteDesignsList}>
        <div className={styles.row}>
            <div className={styles.column}>
                <GridLayout
                    ariaLabel="List of Site Designss."
                    items={items}
                    onRenderGridItem={renderGridItem}
                />
                {seeMore && <div className={styles.seeMore}>
                    {`There are more than ${PREVIEW_ITEMS_COUNT} available Site Designs on your tenant. `}
                    <Link onClick={() => executeAction("GO_TO", { page: "SiteDesignsList" })}>See all Site Designs</Link>
                </div>}
            </div>
        </div>
    </div>;
};