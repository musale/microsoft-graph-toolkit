import { pickerTemplate } from '@microsoft/fast-foundation';
import { pickerStyles, Picker, allComponents } from '@microsoft/fast-components';
import { provideFluentDesignSystem } from '@fluentui/web-components';
import { attr, observable } from '@microsoft/fast-element';
import { IDynamicPerson } from '../../graph/types';
import { DropdownItem } from '../../graph/graph.teams-channels';
import { ValueConverter } from '@microsoft/fast-element';
import { Providers } from '@microsoft/mgt-element';
import { MgtPeoplePicker, MgtTeamsChannelPicker } from '../components';
import { peoplePickerTemplate } from './mgt-picker-template';
// import { MgtGenericPickerBase } from './mgt-fast-base-picker';

/**
 * ensures one call at a time
 *
 * @export
 * @param {*} func
 * @param {*} time
 * @returns
 */
export function debounce(func, time) {
  let timeout;

  return function () {
    const functionCall = () => func.apply(this, arguments);

    clearTimeout(timeout);
    timeout = setTimeout(functionCall, time);
  };
}

/**
 * The entities which can be queried.
 */
export type EntityTypes = 'people' | 'channels' | 'chats';

/**
 * Converts a comma separated string of ints to integer array and back.
 */
export const NumberConverter: ValueConverter = {
  toView: function (value: any) {
    return value.join(',');
  },
  fromView: function (value: any) {
    return value.split(',').map((v: string) => parseInt(v.trim()));
  }
};

/**
 * Converts a comma separated string of strings to an array and back.
 */
export const StringConverter: ValueConverter = {
  toView: function (value: any) {
    return value.join(',');
  },
  fromView: function (value: any) {
    return value.split(',').map((v: string) => v.trim());
  }
};

/**
 * Mgt Generic Picker element.
 */
export class MgtGenericPicker extends Picker {
  constructor() {
    super();
  }

  connectedCallback() {
    super.connectedCallback();
  }

  disconnectedCallback() {
    super.disconnectedCallback();
  }
  public picker: Picker;

  @attr({ attribute: 'entity-types', converter: StringConverter })
  public entityTypes: EntityTypes[];

  @attr({ attribute: 'max-results', converter: NumberConverter })
  public maxResults: number[];

  @observable
  public people: IDynamicPerson[];
  private async peopleChanged(): Promise<void> {
    if (this.$fastController.isConnected) {
      await this.loadState();
    }
  }

  @observable
  public teamsItems: DropdownItem[];

  @observable
  public hasPeople: boolean;

  @observable
  public hasChannels: boolean;

  @observable
  public hasChats: boolean;

  private _debouncedSearch: { (): void; (): void };

  private startSearch = (): void => {
    if (!this._debouncedSearch) {
      this._debouncedSearch = debounce(async () => {
        this.people = [];
        this.teamsItems = [];
        this.showLoading = true;
        await this.loadState();
        this.showLoading = false;
      }, 400);
    }
    this._debouncedSearch();
  };

  private async loadState(): Promise<void> {
    const provider = Providers.globalProvider;
    const entityHasChannels = this.entityTypes.includes('channels');
    const entityHasPeople = this.entityTypes.includes('people');
    const hasChannelScopes = await provider.getAccessTokenForScopes(...MgtTeamsChannelPicker.requiredScopes);
    const hasPeopleScopes = await provider.getAccessTokenForScopes(...MgtPeoplePicker.requiredScopes);

    console.log(this.listItemContentsTemplate);

    console.log({ provider, entityHasChannels, entityHasPeople, hasChannelScopes, hasPeopleScopes });
  }
  public onPickerSelectionChange(e: Event): void {
    if (this.selection === this.picker.selection) {
      return;
    }
    this.selection = this.picker.selection;
  }

  public onPickerQueryChange(e: Event): void {
    this.startSearch();
  }

  public handleMenuOpening(e: Event): void {
    this.showLoading = true;
    this.startSearch();
  }

  public handleMenuClosing(e: Event): void {
    return;
  }
}

/**
 * Composes the <mgt-generic-picker/> HTML element.
 */
export const mgtGenericPicker = MgtGenericPicker.compose({
  baseName: 'generic-picker',
  template: peoplePickerTemplate,
  styles: pickerStyles,
  shadowOptions: {}
});

/**
 * Registers the elements for use in the HTML document.
 */
provideFluentDesignSystem().withPrefix('mgt').register(allComponents, mgtGenericPicker());
