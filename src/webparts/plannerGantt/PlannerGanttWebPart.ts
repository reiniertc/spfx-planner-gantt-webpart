import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneGroup,
  type IPropertyPaneField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneCheckbox,
  PropertyPaneLabel,
  PropertyPaneSlider,
  type IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import * as strings from 'PlannerGanttWebPartStrings';
import PlannerGantt, { MIN_VISIBLE_UNITS, MAX_VISIBLE_UNITS } from './components/PlannerGantt';
import { IPlannerGanttProps } from './components/IPlannerGanttProps';
import { PlannerService } from './services/PlannerService';
import { IPlannerPlanOption, IPlannerBucket, IBucketFilter } from './models/IPlannerModels';

export interface IPlannerGanttWebPartProps {
  planId: string;
  planTitle: string;
  viewMode: string;
  showCompletedTasks: boolean;
  showBucketsAsPhases: boolean;
  colorBarsByStatus: boolean;
  sortTasksByStartDate: boolean;
  sortBucketsByStartDate: boolean;
  showTaskNameOnBar: boolean;
  showCurrentDateLine: boolean;
  showStartDateColumn: boolean;
  showEndDateColumn: boolean;
  showAssigneeColumn: boolean;
  defaultZoomLevel: number;
  showZoomControl: boolean;
  showPrintButton: boolean;
  showAssigneeFilter: boolean;
  bucketFilter?: IBucketFilter;
}

const DEFAULT_ZOOM_LEVEL: number = 3;

const NO_PLAN_SELECTED: string = '';

// Fluent UI's default theme, used until the host page reports its own via onThemeChanged.
const DEFAULT_THEME_PRIMARY: string = '#0078d4';
const DEFAULT_THEME_SECONDARY: string = '#2b88d8';
const DEFAULT_THEME_GREY: string = '#a19f9d';

type LoadStatus = 'idle' | 'loading' | 'loaded' | 'error';

export default class PlannerGanttWebPart extends BaseClientSideWebPart<IPlannerGanttWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _themePrimary: string = DEFAULT_THEME_PRIMARY;
  private _themeSecondary: string = DEFAULT_THEME_SECONDARY;
  private _themeGrey: string = DEFAULT_THEME_GREY;
  private _plannerService: PlannerService;

  private _planOptions: IPropertyPaneDropdownOption[] = [];
  private _plansLoadStatus: LoadStatus = 'idle';

  private _bucketOptions: IPlannerBucket[] = [];
  private _bucketsLoadStatus: LoadStatus = 'idle';
  private _bucketsLoadedForPlanId: string = '';

  public render(): void {
    const element: React.ReactElement<IPlannerGanttProps> = React.createElement(
      PlannerGantt,
      {
        planId: this.properties.planId || NO_PLAN_SELECTED,
        planTitle: this.properties.planTitle,
        viewMode: this.properties.viewMode || 'Week',
        showCompletedTasks: !!this.properties.showCompletedTasks,
        showBucketsAsPhases: this.properties.showBucketsAsPhases !== false,
        colorBarsByStatus: this.properties.colorBarsByStatus !== false,
        sortTasksByStartDate: !!this.properties.sortTasksByStartDate,
        sortBucketsByStartDate: !!this.properties.sortBucketsByStartDate,
        showTaskNameOnBar: this.properties.showTaskNameOnBar !== false,
        showCurrentDateLine: this.properties.showCurrentDateLine !== false,
        showStartDateColumn: this.properties.showStartDateColumn !== false,
        showEndDateColumn: this.properties.showEndDateColumn !== false,
        showAssigneeColumn: this.properties.showAssigneeColumn !== false,
        defaultZoomLevel: this.properties.defaultZoomLevel || DEFAULT_ZOOM_LEVEL,
        showZoomControl: this.properties.showZoomControl !== false,
        showPrintButton: this.properties.showPrintButton !== false,
        showAssigneeFilter: this.properties.showAssigneeFilter !== false,
        // Shallow-copied so a mutated nested property still produces a new
        // reference for React to notice between property pane changes.
        bucketFilter: this.properties.bucketFilter ? { ...this.properties.bucketFilter } : undefined,
        isDarkTheme: this._isDarkTheme,
        themePrimary: this._themePrimary,
        themeSecondary: this._themeSecondary,
        themeGrey: this._themeGrey,
        plannerService: this._plannerService
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._plannerService = new PlannerService(
      (): Promise<MSGraphClientV3> => this.context.msGraphClientFactory.getClient('3')
    );
    return Promise.resolve();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors, palette } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

    if (palette) {
      this._themePrimary = palette.themePrimary || DEFAULT_THEME_PRIMARY;
      this._themeSecondary = palette.themeSecondary || DEFAULT_THEME_SECONDARY;
      this._themeGrey = palette.neutralTertiary || DEFAULT_THEME_GREY;
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart(): void {
    if (this._plansLoadStatus === 'idle') {
      this._loadPlanOptions();
    }
    this._loadBucketOptionsIfNeeded(this.properties.planId);
  }

  private _loadPlanOptions(): void {
    this._plansLoadStatus = 'loading';
    this._planOptions = [];
    this.context.propertyPane.refresh();

    this._plannerService
      .getAvailablePlans()
      .then((plans: IPlannerPlanOption[]) => {
        this._planOptions = plans.map(plan => ({
          key: plan.planId,
          text: `${plan.planTitle} (${plan.groupName})`
        }));
        this._plansLoadStatus = 'loaded';
        this.context.propertyPane.refresh();
      })
      .catch(() => {
        this._plansLoadStatus = 'error';
        this.context.propertyPane.refresh();
      });
  }

  private _loadBucketOptionsIfNeeded(planId: string): void {
    if (!planId) {
      this._bucketOptions = [];
      this._bucketsLoadStatus = 'idle';
      this._bucketsLoadedForPlanId = '';
      return;
    }

    if (this._bucketsLoadedForPlanId === planId && this._bucketsLoadStatus !== 'idle') {
      return;
    }

    this._bucketsLoadStatus = 'loading';
    this._bucketsLoadedForPlanId = planId;
    this._bucketOptions = [];
    this.context.propertyPane.refresh();

    this._plannerService
      .getBuckets(planId)
      .then((buckets: IPlannerBucket[]) => {
        if (this._bucketsLoadedForPlanId !== planId) {
          return; // the user picked a different plan while this call was in flight
        }
        this._bucketOptions = buckets;
        this._bucketsLoadStatus = 'loaded';

        // First time buckets are seen for this plan: default every bucket to visible.
        if (!this.properties.bucketFilter) {
          const defaultFilter: IBucketFilter = {};
          buckets.forEach(bucket => {
            defaultFilter[bucket.id] = true;
          });
          this.properties.bucketFilter = defaultFilter;
        }

        this.context.propertyPane.refresh();
      })
      .catch(() => {
        this._bucketsLoadStatus = 'error';
        this.context.propertyPane.refresh();
      });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    if (propertyPath === 'planId') {
      const selected: IPropertyPaneDropdownOption | undefined = this._planOptions.filter(option => option.key === newValue)[0];
      this.properties.planTitle = selected ? selected.text : '';
      this.properties.bucketFilter = undefined;
      this._loadBucketOptionsIfNeeded(newValue as string);
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const isLoadingPlans: boolean = this._plansLoadStatus === 'loading' || this._plansLoadStatus === 'idle';
    const plansHaveError: boolean = this._plansLoadStatus === 'error';
    let planDropdownOptions: IPropertyPaneDropdownOption[] = this._planOptions;
    if (isLoadingPlans) {
      planDropdownOptions = [{ key: NO_PLAN_SELECTED, text: strings.LoadingPlansLabel }];
    } else if (plansHaveError) {
      planDropdownOptions = [{ key: NO_PLAN_SELECTED, text: strings.ErrorLoadingPlansLabel }];
    }

    const groups: IPropertyPaneGroup[] = [
      {
        groupName: strings.BasicGroupName,
        groupFields: [
          PropertyPaneDropdown('planId', {
            label: strings.PlanFieldLabel,
            options: planDropdownOptions,
            disabled: isLoadingPlans || plansHaveError,
            selectedKey: this.properties.planId
          })
        ]
      },
      {
        groupName: strings.DisplayGroupName,
        groupFields: [
          PropertyPaneDropdown('viewMode', {
            label: strings.ViewModeFieldLabel,
            options: [
              { key: 'Day', text: strings.ViewModeDay },
              { key: 'Week', text: strings.ViewModeWeek },
              { key: 'Month', text: strings.ViewModeMonth }
            ],
            selectedKey: this.properties.viewMode || 'Week'
          }),
          PropertyPaneToggle('showCompletedTasks', {
            label: strings.ShowCompletedFieldLabel
          }),
          PropertyPaneToggle('showBucketsAsPhases', {
            label: strings.ShowBucketsAsPhasesFieldLabel,
            checked: this.properties.showBucketsAsPhases !== false
          }),
          PropertyPaneToggle('sortBucketsByStartDate', {
            label: strings.SortBucketsByStartDateFieldLabel,
            onText: strings.SortBucketsByStartDateOnText,
            offText: strings.SortBucketsByStartDateOffText,
            checked: !!this.properties.sortBucketsByStartDate
          }),
          PropertyPaneToggle('sortTasksByStartDate', {
            label: strings.SortTasksByStartDateFieldLabel,
            checked: !!this.properties.sortTasksByStartDate
          }),
          PropertyPaneToggle('colorBarsByStatus', {
            label: strings.ColorBarsByStatusFieldLabel,
            checked: this.properties.colorBarsByStatus !== false
          }),
          PropertyPaneToggle('showTaskNameOnBar', {
            label: strings.ShowTaskNameOnBarFieldLabel,
            checked: this.properties.showTaskNameOnBar !== false
          }),
          PropertyPaneToggle('showCurrentDateLine', {
            label: strings.ShowCurrentDateLineFieldLabel,
            checked: this.properties.showCurrentDateLine !== false
          })
        ]
      },
      {
        groupName: strings.ColumnsGroupName,
        groupFields: [
          PropertyPaneToggle('showStartDateColumn', {
            label: strings.ShowStartDateColumnFieldLabel,
            checked: this.properties.showStartDateColumn !== false
          }),
          PropertyPaneToggle('showEndDateColumn', {
            label: strings.ShowEndDateColumnFieldLabel,
            checked: this.properties.showEndDateColumn !== false
          }),
          PropertyPaneToggle('showAssigneeColumn', {
            label: strings.ShowAssigneeColumnFieldLabel,
            checked: this.properties.showAssigneeColumn !== false
          })
        ]
      },
      {
        groupName: strings.ToolbarGroupName,
        groupFields: [
          PropertyPaneSlider('defaultZoomLevel', {
            label: strings.DefaultZoomLevelFieldLabel,
            min: MIN_VISIBLE_UNITS,
            max: MAX_VISIBLE_UNITS,
            step: 1,
            value: this.properties.defaultZoomLevel || DEFAULT_ZOOM_LEVEL
          }),
          PropertyPaneToggle('showZoomControl', {
            label: strings.ShowZoomControlFieldLabel,
            checked: this.properties.showZoomControl !== false
          }),
          PropertyPaneToggle('showAssigneeFilter', {
            label: strings.ShowAssigneeFilterFieldLabel,
            checked: this.properties.showAssigneeFilter !== false
          }),
          PropertyPaneToggle('showPrintButton', {
            label: strings.ShowPrintButtonFieldLabel,
            checked: this.properties.showPrintButton !== false
          })
        ]
      }
    ];

    if (this.properties.planId) {
      groups.push({
        groupName: strings.BucketsGroupName,
        groupFields: this._getBucketFilterFields()
      });
    }

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups
        }
      ]
    };
  }

  private _getBucketFilterFields(): IPropertyPaneField<unknown>[] {
    if (this._bucketsLoadStatus === 'loading' || this._bucketsLoadStatus === 'idle') {
      return [PropertyPaneLabel('bucketsStatus', { text: strings.LoadingBucketsLabel })];
    }

    if (this._bucketsLoadStatus === 'error') {
      return [PropertyPaneLabel('bucketsStatus', { text: strings.ErrorLoadingBucketsLabel })];
    }

    if (this._bucketOptions.length === 0) {
      return [PropertyPaneLabel('bucketsStatus', { text: strings.NoBucketsLabel })];
    }

    return this._bucketOptions.map(bucket =>
      PropertyPaneCheckbox(`bucketFilter.${bucket.id}`, {
        text: bucket.name,
        checked: this.properties.bucketFilter ? this.properties.bucketFilter[bucket.id] !== false : true
      })
    );
  }
}
