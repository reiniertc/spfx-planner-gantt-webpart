import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  type IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import * as strings from 'PlannerGanttWebPartStrings';
import PlannerGantt from './components/PlannerGantt';
import { IPlannerGanttProps } from './components/IPlannerGanttProps';
import { PlannerService } from './services/PlannerService';
import { IPlannerPlanOption } from './models/IPlannerModels';

export interface IPlannerGanttWebPartProps {
  planId: string;
  planTitle: string;
  viewMode: string;
  showCompletedTasks: boolean;
}

const NO_PLAN_SELECTED: string = '';

export default class PlannerGanttWebPart extends BaseClientSideWebPart<IPlannerGanttWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _plannerService: PlannerService;
  private _planOptions: IPropertyPaneDropdownOption[] = [];
  private _plansLoadStatus: 'idle' | 'loading' | 'loaded' | 'error' = 'idle';

  public render(): void {
    const element: React.ReactElement<IPlannerGanttProps> = React.createElement(
      PlannerGantt,
      {
        planId: this.properties.planId || NO_PLAN_SELECTED,
        planTitle: this.properties.planTitle,
        viewMode: this.properties.viewMode || 'Week',
        showCompletedTasks: !!this.properties.showCompletedTasks,
        isDarkTheme: this._isDarkTheme,
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
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
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

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    if (propertyPath === 'planId') {
      const selected: IPropertyPaneDropdownOption | undefined = this._planOptions.filter(option => option.key === newValue)[0];
      this.properties.planTitle = selected ? selected.text : '';
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const isLoadingPlans: boolean = this._plansLoadStatus === 'loading' || this._plansLoadStatus === 'idle';
    const hasError: boolean = this._plansLoadStatus === 'error';
    let planDropdownOptions: IPropertyPaneDropdownOption[] = this._planOptions;
    if (isLoadingPlans) {
      planDropdownOptions = [{ key: NO_PLAN_SELECTED, text: strings.LoadingPlansLabel }];
    } else if (hasError) {
      planDropdownOptions = [{ key: NO_PLAN_SELECTED, text: strings.ErrorLoadingPlansLabel }];
    }

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneDropdown('planId', {
                  label: strings.PlanFieldLabel,
                  options: planDropdownOptions,
                  disabled: isLoadingPlans || hasError,
                  selectedKey: this.properties.planId
                }),
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
