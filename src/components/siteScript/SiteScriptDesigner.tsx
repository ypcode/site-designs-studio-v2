import * as React from "react";
import { useAppContext } from "../../app/App";
import { IApplicationState } from "../../app/ApplicationState";
import { ISiteScriptContent, ISiteScriptAction } from "../../models/ISiteScript";
import { SiteScriptSchemaServiceKey } from "../../services/siteScriptSchema/SiteScriptSchemaService";
import { ActionType } from "../../app/IApplicationAction";
import { GenericObjectEditor } from "../common/genericObjectEditor/GenericObjectEditor";
import styles from "./SiteScriptDesigner.module.scss";
import { Adder, IAddableItem } from "../common/Adder/Adder";
import { IconButton, getId, Link, Label } from "office-ui-fabric-react";
import { useState, useEffect } from "react";
import { isEqual } from "@microsoft/sp-lodash-subset";
import { getTrimmedText } from "../../utils/textUtils";
import { useConstCallback } from "@uifabric/react-hooks";


interface IEditingActionIndexes {
    actionIndex: number;
    subActionIndex: number;
}

export interface ISiteScriptActionDesignerBlockProps {
    actionKey: string;
    siteScriptAction: ISiteScriptAction;
    parentSiteScriptAction?: ISiteScriptAction;
    isEditing: boolean;
    isEditingSubActionIndex?: number;
    onSiteScriptActionUpdated: (key: string, updatedScriptAction: ISiteScriptAction) => void;
    onSiteScriptActionRemoved: (key: string) => void;
    onEditingChanged?: (isEditing: boolean, subActionIndex?: number) => void;
}

const SEE_PROPERTIES_DEFAULT_COUNT = 2;
const SUMMARY_VALUE_MAX_LEN = 60;

export const SiteScriptActionDesignerBlock = (props: ISiteScriptActionDesignerBlockProps) => {

    const [appContext] = useAppContext<IApplicationState, ActionType>();

    // Get service references
    const siteScriptSchemaService = appContext.serviceScope.consume(SiteScriptSchemaServiceKey);

    const actionSchema = props.parentSiteScriptAction ?
        siteScriptSchemaService.getSubActionSchema(props.parentSiteScriptAction, props.siteScriptAction)
        : siteScriptSchemaService.getActionSchema(props.siteScriptAction);

    const [seeMore, setSeeMore] = useState<boolean>(false);

    const newSubActionKey = () => getId(`ScriptAction_${props.actionKey}_`);
    const getKeyedSubActionsCollection = (scriptActions: ISiteScriptAction[]) => scriptActions.map(sa => ({ key: newSubActionKey(), action: sa }));
    const [keyedSubActions, setKeyedSubActions] = useState<{ key: string; action: ISiteScriptAction }[]>(props.siteScriptAction.subactions
        ? getKeyedSubActionsCollection(props.siteScriptAction.subactions)
        : []);

    const getActionDescription = (() => siteScriptSchemaService.getDescriptionFromActionSchema(actionSchema));
    const getActionLabel = (() => siteScriptSchemaService.getLabelFromActionSchema(actionSchema));

    const onSubActionUpdated = ((key: string, siteScriptAction: ISiteScriptAction) => {
        const updatedSubActions = keyedSubActions.map((keyedAction) => keyedAction.key == key ? { key, action: siteScriptAction } : keyedAction);
        setKeyedSubActions(updatedSubActions);
        const updated = {
            ...props.siteScriptAction,
            subactions: updatedSubActions.map(sa => sa.action)
        };
        props.onSiteScriptActionUpdated(props.actionKey, updated);
    });

    const onSubActionRemoved = ((key: string) => {
        const updatedSubActions = keyedSubActions.filter((keyedAction) => keyedAction.key != key);
        setKeyedSubActions(updatedSubActions);
        const updated = {
            ...props.siteScriptAction,
            subactions: updatedSubActions.map(sa => sa.action)
        };
        props.onSiteScriptActionUpdated(props.actionKey, updated);
    });

    const onSubActionAdded = ((verb: string) => {
        const newAction = siteScriptSchemaService.getNewActionFromVerb(verb);
        const updatedSubActions = [...keyedSubActions || [], { key: newSubActionKey(), action: newAction }];
        setKeyedSubActions(updatedSubActions);
        const updated = {
            ...props.siteScriptAction,
            subactions: updatedSubActions.map(sa => sa.action)
        };
        props.onSiteScriptActionUpdated(props.actionKey, updated);
    });

    const getAddableActions = (() => {
        const groupLabel = `${getActionLabel()} - Subactions`;
        return {
            [groupLabel]: siteScriptSchemaService.getAvailableSubActions(props.siteScriptAction).map(a => ({
                group: groupLabel,
                iconName: "SetAction",
                key: a.verb,
                text: a.label,
                item: a
            } as IAddableItem))
        };
    });

    const toggleEdit = (() => {
        if (props.onEditingChanged) {
            props.onEditingChanged(!props.isEditing);
        }
    });

    const onSubActionEditingChanged = ((isEditing: boolean, subActionIndex: number) => {
        if (props.onEditingChanged) {
            props.onEditingChanged(isEditing, subActionIndex);
        }
    });

    const getPropertiesAndValues = (() => {
        return Object.keys(actionSchema.properties).filter(p => p != "verb").map(p => ({
            title: actionSchema.properties[p].title,
            value: props.siteScriptAction[p]
        }));
    });

    const renderSummaryContent = (() => {
        const summaryValues = getPropertiesAndValues();
        if (!seeMore) {
            const previewSummary = summaryValues.slice(0, SEE_PROPERTIES_DEFAULT_COUNT);
            const displaySeeMoreLink = summaryValues.length >= SEE_PROPERTIES_DEFAULT_COUNT && !seeMore;
            return <ul>
                {previewSummary.map((pv, index) => <li key={`${props.actionKey}_prop_${index}`}>{pv.title}: <strong>{(!pv.value && pv.value !== false) ? "Not set" : getTrimmedText(pv.value, SUMMARY_VALUE_MAX_LEN)}</strong></li>)}
                {displaySeeMoreLink && <li key={`${props.actionKey}_more_prop`}><Link onClick={() => setSeeMore(true)}>...</Link></li>}
            </ul>;
        } else {
            return <ul>
                {summaryValues.map((pv, index) => <li key={`${props.actionKey}_prop_${index}`}>{pv.title}: <strong>{!pv.value ? "Not set" : pv.value}</strong></li>)}
            </ul>;
        }
    });


    return <div className={`${styles.siteScriptAction} ${props.isEditing ? styles.isEditing : ""}`}>
        <h4 title={getActionDescription()}>
            {getActionLabel()}
        </h4>
        <div className={styles.tools}>
            <IconButton iconProps={{ iconName: props.isEditing ? "Accept" : "Edit" }} onClick={() => toggleEdit()} />
            {!props.isEditing && <IconButton iconProps={{ iconName: "Delete" }} onClick={() => props.onSiteScriptActionRemoved(props.actionKey)} />}
        </div>
        <div className={`${styles.summary} ${props.isEditing ? styles.isEditing : styles.isNotEditing}`}>
            {renderSummaryContent()}
        </div>
        <div className={styles.properties}>
            <GenericObjectEditor
                ignoredProperties={["verb"]}
                object={props.siteScriptAction}
                schema={actionSchema}
                customRenderers={{
                    "subactions": () => <div className={styles.subactions}>
                        <Label>Subactions</Label>
                        {keyedSubActions.map((keyedSubAction, index) => <SiteScriptActionDesignerBlock key={keyedSubAction.key} actionKey={keyedSubAction.key}
                            parentSiteScriptAction={props.siteScriptAction}
                            siteScriptAction={keyedSubAction.action}
                            onSiteScriptActionUpdated={onSubActionUpdated}
                            onSiteScriptActionRemoved={onSubActionRemoved}
                            isEditing={props.isEditing && props.isEditingSubActionIndex == index}
                            onEditingChanged={(isEditing) => onSubActionEditingChanged(isEditing, index)} />)}
                        <Adder items={getAddableActions()}
                            searchBoxPlaceholderText="Search a sub action..."
                            onSelectedItem={(item) => onSubActionAdded(item.key)} />
                    </div>
                }}
                onObjectChanged={(o) => props.onSiteScriptActionUpdated(props.actionKey, o as ISiteScriptAction)} >
            </GenericObjectEditor>
        </div>
    </div>;
};

export interface ISiteScriptDesignerProps {
    siteScriptContent: ISiteScriptContent;
    onSiteScriptContentUpdated: (updatedSiteScriptContent: ISiteScriptContent) => void;
}

export const SiteScriptDesigner = (props: ISiteScriptDesignerProps) => {
    const [appContext] = useAppContext<IApplicationState, ActionType>();

    // Get service references
    const siteScriptSchemaService = appContext.serviceScope.consume(SiteScriptSchemaServiceKey);

    const [editingIndexes, setEditingIndexes] = useState<IEditingActionIndexes>({ actionIndex: null, subActionIndex: null });

    const newActionKey = () => getId("ScriptAction_");
    const getKeyedActionsCollection = (scriptActions: ISiteScriptAction[]) => scriptActions.map(sa => ({ key: newActionKey(), action: sa }));

    const [keyedActions, setKeyedActions] = useState<{ key: string; action: ISiteScriptAction }[]>(getKeyedActionsCollection((props.siteScriptContent && props.siteScriptContent.actions) || []));

    useEffect(() => {
        console.log("Site script content changed");
        const actionsFromState = keyedActions.map(ka => ka.action);
        // Refresh the actions if updated from outsite
        if (props.siteScriptContent && !isEqual(props.siteScriptContent.actions, actionsFromState)) {
            setKeyedActions(getKeyedActionsCollection(props.siteScriptContent.actions));
            setEditingIndexes({ actionIndex: null, subActionIndex: null });
        }
    }, [props.siteScriptContent]);

    const onActionUpdated = ((key: string, siteScriptAction: ISiteScriptAction) => {
        const updatedActions = keyedActions.map((keyedAction) => key == keyedAction.key ? { key, action: siteScriptAction } : keyedAction);
        setKeyedActions(updatedActions);
        const updated = {
            ...props.siteScriptContent,
            actions: updatedActions.map(ka => ka.action)
        };
        props.onSiteScriptContentUpdated(updated);
    });

    const onActionAdded = ((verb: string) => {
        const newAction = siteScriptSchemaService.getNewActionFromVerb(verb);
        const updatedActions = [...keyedActions, { key: newActionKey(), action: newAction }];
        setKeyedActions(updatedActions);
        const updated = {
            ...props.siteScriptContent,
            actions: updatedActions.map(ka => ka.action)
        };
        props.onSiteScriptContentUpdated(updated);
    });

    const onActionRemoved = ((key: string) => {
        const updatedActions = keyedActions.filter((keyedAction) => key != keyedAction.key);
        setKeyedActions(updatedActions);
        const updated = {
            ...props.siteScriptContent,
            actions: updatedActions.map(ka => ka.action)
        };
        // Clear all editions where an action is removed
        setEditingIndexes({ actionIndex: null, subActionIndex: null });
        props.onSiteScriptContentUpdated(updated);
    });

    const onEditingChanged = ((isEditing: boolean, actionIndex: number, subActionIndex?: number) => {
        if (isEditing) {
            setEditingIndexes({ actionIndex, subActionIndex });
        } else {
            if (editingIndexes.actionIndex == actionIndex) {
                if (subActionIndex == editingIndexes.subActionIndex) {
                    setEditingIndexes({ actionIndex, subActionIndex: null });
                } else {
                    setEditingIndexes({ actionIndex: null, subActionIndex: null });
                }
            }
        }
    });

    const getAddableActions = useConstCallback(() => {
        return {
            "Actions": siteScriptSchemaService.getAvailableActions().map(a => ({
                group: "Actions",
                iconName: "SetAction",
                key: a.verb,
                text: a.label,
                item: a
            } as IAddableItem))
        };
    });

    // TODO Implement orderable collection
    return <div className={styles.SiteScriptDesigner}>
        {keyedActions.map((keyedAction, index) => <SiteScriptActionDesignerBlock key={keyedAction.key} actionKey={keyedAction.key}
            siteScriptAction={keyedAction.action}
            onSiteScriptActionUpdated={onActionUpdated}
            onSiteScriptActionRemoved={onActionRemoved}
            isEditing={index == editingIndexes.actionIndex}
            onEditingChanged={(isEditing, subActionIndex) => onEditingChanged(isEditing, index, subActionIndex)}
            isEditingSubActionIndex={editingIndexes.subActionIndex} />)}
        <Adder items={getAddableActions()}
            searchBoxPlaceholderText="Search an action..."
            onSelectedItem={(item) => onActionAdded(item.key)} />
    </div>;
};