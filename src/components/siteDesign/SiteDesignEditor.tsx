import * as React from "react";
import { SortableContainer, SortableHandle, SortableElement } from 'react-sortable-hoc';
import { ISiteDesign, WebTemplate } from "../../models/ISiteDesign";
import { useState, useEffect } from "react";
import styles from "./SiteDesignEditor.module.scss";
import { TextField, Dropdown, Label, ActionButton, PrimaryButton, DocumentCardPreview, ImageFit, IDocumentCardPreviewProps, Spinner, SpinnerType, Stack, Toggle, DefaultButton, ProgressIndicator } from "office-ui-fabric-react";
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { useAppContext } from "../../app/App";
import { IApplicationState } from "../../app/ApplicationState";
import { ActionType, ISetAllAvailableSiteDesigns, IGoToActionArgs } from "../../app/IApplicationAction";
import { Adder, IAddableItem } from "../common/Adder/Adder";
import { SiteDesignsServiceKey } from "../../services/siteDesigns/SiteDesignsService";
import { ISiteScript } from "../../models/ISiteScript";
import { find } from "@microsoft/sp-lodash-subset";
import { Confirm } from "../common/Confirm/Confirm";

export interface ISiteDesignEditorProps {
    siteDesign: ISiteDesign;
}

export interface ISiteDesignAssociatedSiteScriptsProps extends ISiteDesignEditorProps {
    onAssociatedSiteScriptAdded: (siteScriptId: string) => void;
    onAssociatedSiteScriptRemoved: (siteScriptId: string) => void;
    onAssociatedSiteScriptsReordered: (reordered: string[]) => void;
}

interface ISortEndEventArgs {
    oldIndex: number;
    newIndex: number;
    collection: any[];
}

export const SiteDesignAssociatedScripts = (props: ISiteDesignAssociatedSiteScriptsProps) => {
    const [appContext, execute] = useAppContext<IApplicationState, ActionType>();

    if (!appContext.allAvailableSiteScripts || appContext.allAvailableSiteScripts.length == 0) {
        return <div>Loading site scripts...</div>;
    }

    const selectedSiteScripts = props.siteDesign.SiteScriptIds.map(id => find(appContext.allAvailableSiteScripts, ss => ss.Id == id));
    const adderItems: IAddableItem[] = appContext.allAvailableSiteScripts.filter(ss => props.siteDesign.SiteScriptIds.indexOf(ss.Id) < 0).map(item => ({
        iconName: "Script",
        group: null,
        key: item.Id,
        text: item.Title,
        item
    }));

    const renderSiteScriptItem = (siteScript: ISiteScript, index: number) => {
        const DragHandle = SortableHandle(() => (
            <div className={styles.column10}>
                <h4>{siteScript.Title}</h4>
                <div>{siteScript.Description}</div>
            </div>
        ));

        const SortableItem = SortableElement(({ value }) => <div key={`selectedSiteScript_${index}`} className={styles.selectedSiteScript}>
            <div className={styles.row}>
                <DragHandle />
                <div className={`${styles.column2} ${styles.righted}`}>
                    <ActionButton iconProps={{ iconName: "Delete" }} onClick={() => props.onAssociatedSiteScriptRemoved(value.Id)} />
                </div>
            </div>
        </div>);

        return <SortableItem key={`SiteScript_${siteScript.Id}`} value={siteScript} index={index} />;
    };

    const SortableListContainer = SortableContainer(({ items }) => {
        return <div>{items.map(renderSiteScriptItem)}</div>;
    });

    const onSortChanged = (args: ISortEndEventArgs) => {
        const toSortSiteScriptIds = [...props.siteDesign.SiteScriptIds];
        toSortSiteScriptIds.splice(args.oldIndex, 1);
        toSortSiteScriptIds.splice(args.newIndex, 0, props.siteDesign.SiteScriptIds[args.oldIndex]);
        props.onAssociatedSiteScriptsReordered(toSortSiteScriptIds);
    };

    return <>
        <Label>Associated Site Scripts: </Label>
        <SortableListContainer
            items={selectedSiteScripts}
            // onSortStart={(args) => this._onSortStart(args)}
            onSortEnd={(args: any) => onSortChanged(args)}
            lockToContainerEdges={true}
            useDragHandle={true}
        />
        <Adder items={{ "Available Site Scripts": adderItems }}
            onSelectedItem={(item) => props.onAssociatedSiteScriptAdded(item.item.Id)} />
    </>;
};

export const SiteDesignEditor = (props: ISiteDesignEditorProps) => {

    const [appContext, execute] = useAppContext<IApplicationState, ActionType>();
    const siteDesignsService = appContext.serviceScope.consume(SiteDesignsServiceKey);

    const [editingSiteDesign, setEditingSiteDesign] = useState<ISiteDesign>(props.siteDesign);
    const [focusedField, setFocusedField] = useState<string>("");
    const [isSaving, setIsSaving] = useState<boolean>(false);
    const [chosenFile, setChosenFile] = useState<IFilePickerResult>({
        fileName: "",
        fileAbsoluteUrl: "",
        fileNameWithoutExtension: "",
        spItemUrl: "",
        downloadFileContent: null
    });

    const setLoading = (loading: boolean) => {
        execute("SET_LOADING", { loading });
    };

    useEffect(() => {
        if (!props.siteDesign.Id) {
            setEditingSiteDesign(props.siteDesign);
            return;
        }

        setLoading(true);
        siteDesignsService.getSiteDesign(props.siteDesign.Id).then(siteDesign => {
            setEditingSiteDesign(siteDesign);
        }).catch(error => {
            console.error(`The Site Design ${props.siteDesign.Id} could not be loaded`, error);
        }).then(() => {
            setLoading(false);
        });
    }, [props.siteDesign]);

    const onTitleChanged = (ev: any, title: string) => {
        setEditingSiteDesign({ ...editingSiteDesign, Title: title });
    };

    const onDescriptionChanged = (ev: any, description: string) => {
        setEditingSiteDesign({ ...editingSiteDesign, Description: description });
    };

    const onIsDefaultChanged = (ev: any, isDefault: boolean) => {
        setEditingSiteDesign({ ...editingSiteDesign, IsDefault: isDefault });
    };

    const onPreviewImageAltTextChanged = (ev: any, previewImageAltText: string) => {
        setEditingSiteDesign({ ...editingSiteDesign, PreviewImageAltText: previewImageAltText });
    };

    const onPreviewImageRemoved = () => {
        setEditingSiteDesign({ ...editingSiteDesign, PreviewImageUrl: null });
    };

    const onVersionChanged = (ev: any, version: string) => {
        const versionInt = parseInt(version);
        if (!isNaN(versionInt)) {
            setEditingSiteDesign({ ...editingSiteDesign, Version: versionInt });
        }
    };

    const onWebTemplateChanged = (webTemplate: string) => {
        setEditingSiteDesign({ ...editingSiteDesign, WebTemplate: webTemplate });
    };

    const onAssociatedSiteScriptsAdded = (siteScriptId: string) => {
        if (!editingSiteDesign.SiteScriptIds) {
            return;
        }

        const newSiteScriptIds = editingSiteDesign.SiteScriptIds.filter(sid => sid != siteScriptId).concat(siteScriptId);
        setEditingSiteDesign({ ...editingSiteDesign, SiteScriptIds: newSiteScriptIds });
    };


    const onAssociatedSiteScriptsRemoved = (siteScriptId: string) => {
        if (!editingSiteDesign.SiteScriptIds) {
            return;
        }

        const newSiteScriptIds = editingSiteDesign.SiteScriptIds.filter(sid => sid != siteScriptId);
        setEditingSiteDesign({ ...editingSiteDesign, SiteScriptIds: newSiteScriptIds });
    };

    const onAssociatedSiteScriptsReordered = (reorderedSiteScriptIds: string[]) => {
        if (!editingSiteDesign.SiteScriptIds) {
            return;
        }

        setEditingSiteDesign({ ...editingSiteDesign, SiteScriptIds: reorderedSiteScriptIds });
    };

    const onSave = async () => {
        setIsSaving(true);
        await siteDesignsService.saveSiteDesign(editingSiteDesign);
        const refreshedSiteDesigns = await siteDesignsService.getSiteDesigns();
        execute("SET_ALL_AVAILABLE_SITE_DESIGNS", { siteDesigns: refreshedSiteDesigns } as ISetAllAvailableSiteDesigns);
        setIsSaving(false);
    };

    const onDelete = async () => {

        if (!await Confirm.show({
            title: `Delete Site Design`,
            message: `Are you sure you want to delete ${editingSiteDesign.Title || "this Site Design"} ?`
        })) {
            return;
        }

        setIsSaving(true);
        await siteDesignsService.deleteSiteDesign(editingSiteDesign);
        const refreshedSiteDesigns = await siteDesignsService.getSiteDesigns();
        execute("SET_ALL_AVAILABLE_SITE_DESIGNS", { siteDesigns: refreshedSiteDesigns } as ISetAllAvailableSiteDesigns);
        execute("GO_TO", { page: "SiteDesignsList" } as IGoToActionArgs);
        setIsSaving(false);
    };

    const previewProps: IDocumentCardPreviewProps = {
        previewImages: [
            {
                previewImageSrc: editingSiteDesign.PreviewImageUrl,
                imageFit: ImageFit.centerContain,
                height: 300
            }
        ]
    };

    const isLoading = appContext.isLoading;
    return <div className={styles.SiteDesignEditor}>
        <div className={styles.row}>
            <div className={styles.columnLayout}>
                <div className={styles.row}>
                    <div className={styles.column11}>
                        <TextField
                            onFocus={() => setFocusedField("Title")}
                            onBlur={() => setFocusedField(null)}
                            styles={{
                                field: {
                                    fontSize: "32px",
                                    lineHeight: "45px",
                                    height: "45px"
                                },
                                root: {
                                    height: "60px",
                                    marginTop: "5px",
                                    marginBottom: "5px"
                                }
                            }}
                            placeholder="Enter the name of the Site Design..."
                            borderless
                            readOnly={isLoading}
                            value={editingSiteDesign.Title}
                            onChange={onTitleChanged} />
                        {isLoading && <ProgressIndicator />}
                    </div>
                    {!isLoading && <div className={`${styles.column1} ${styles.righted}`}>
                        <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 15 }}>
                            <DefaultButton disabled={isSaving} text="Delete" iconProps={{ iconName: "Delete" }} onClick={() => onDelete()} />
                            <PrimaryButton disabled={isSaving} text="Save" iconProps={{ iconName: "Save" }} onClick={() => onSave()} />
                        </Stack>
                    </div>}
                </div>
                {!isLoading && <div className={styles.row}>
                    <div className={styles.half}>
                        {editingSiteDesign.Id && <div className={styles.row}>
                            <div className={styles.column6}>
                                <TextField
                                    label="Id"
                                    readOnly
                                    value={editingSiteDesign.Id} />
                            </div>
                            <div className={styles.column4}>
                                <Dropdown
                                    label="Site Template"
                                    options={[
                                        { key: WebTemplate.TeamSite.toString(), text: 'Team Site' },
                                        { key: WebTemplate.CommunicationSite.toString(), text: 'Communication Site' }
                                    ]}
                                    onFocus={() => setFocusedField("WebTemplate")}
                                    onBlur={() => setFocusedField(null)}
                                    selectedKey={editingSiteDesign.WebTemplate}
                                    onChanged={(v) => onWebTemplateChanged(v.key as string)}
                                />
                            </div>
                            <div className={styles.column2}>
                                <TextField
                                    label="Version"
                                    onFocus={() => setFocusedField("Version")}
                                    onBlur={() => setFocusedField(null)}
                                    value={editingSiteDesign.Version.toString()}
                                    onChange={onVersionChanged} />
                            </div>
                        </div>}
                        {!editingSiteDesign.Id && <div className={styles.row}>
                            <div className={styles.column}>
                                <Dropdown
                                    label="Site Template"
                                    options={[
                                        { key: WebTemplate.TeamSite.toString(), text: 'Team Site' },
                                        { key: WebTemplate.CommunicationSite.toString(), text: 'Communication Site' }
                                    ]}
                                    onFocus={() => setFocusedField("WebTemplate")}
                                    onBlur={() => setFocusedField(null)}
                                    selectedKey={editingSiteDesign.WebTemplate}
                                    onChanged={(v) => onWebTemplateChanged(v.key as string)}
                                />
                            </div>
                        </div>}
                        <div className={styles.row}>
                            <div className={styles.column}>
                                <Toggle
                                    label="Is Default ?"
                                    checked={editingSiteDesign.IsDefault}
                                    onChange={onIsDefaultChanged}
                                />
                            </div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column}>
                                <TextField
                                    label="Description"
                                    value={editingSiteDesign.Description}
                                    multiline={true}
                                    borderless
                                    rows={5}
                                    onChange={onDescriptionChanged}
                                />
                            </div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column}>
                                <SiteDesignAssociatedScripts siteDesign={editingSiteDesign}
                                    onAssociatedSiteScriptAdded={onAssociatedSiteScriptsAdded}
                                    onAssociatedSiteScriptRemoved={onAssociatedSiteScriptsRemoved}
                                    onAssociatedSiteScriptsReordered={onAssociatedSiteScriptsReordered} />
                            </div>
                        </div>
                    </div>
                    <div className={styles.siteDesignImage}>
                        <div className={styles.righted}>
                            <Stack horizontal horizontalAlign="end">
                                {editingSiteDesign.PreviewImageUrl && <ActionButton text="Remove preview image" iconProps={{ iconName: "Delete" }} onClick={onPreviewImageRemoved} />}
                                <FilePicker
                                    accepts={[".gif", ".jpg", ".jpeg", ".bmp", ".png"]}
                                    buttonIcon="FileImage"
                                    onSave={setChosenFile}
                                    onChanged={setChosenFile}
                                    buttonLabel={editingSiteDesign.PreviewImageUrl ? "Modify preview image" : "Add a preview image"}
                                    context={appContext.componentContext}
                                />
                            </Stack>
                        </div>
                        {!editingSiteDesign.PreviewImageUrl && <div><Placeholder iconName='FileImage'
                            iconText='No preview image...'
                            description='There is no defined preview image for this Site Design...' />
                        </div>}
                        {editingSiteDesign.PreviewImageUrl && <div className={styles.imgPlaceholder}>
                            <DocumentCardPreview {...previewProps} />
                        </div>}
                        <div>
                            {editingSiteDesign.PreviewImageUrl && <TextField
                                value={editingSiteDesign.PreviewImageAltText}
                                borderless
                                onFocus={() => setFocusedField("PreviewImageAltText")}
                                onBlur={() => setFocusedField(null)}
                                styles={{
                                    field: {
                                        fontSize: "16px",
                                        lineHeight: "30px",
                                        height: "30px",
                                        textAlign: "center"
                                    },
                                    root: {
                                        height: "30px",
                                        width: "80%",
                                        margin: "auto",
                                        marginTop: "5px"
                                    }
                                }}
                                placeholder="Enter the alternative text for preview image..."
                                onChange={onPreviewImageAltTextChanged}
                            />}
                        </div>

                    </div>
                </div>}
            </div>
        </div>
    </div>;
};