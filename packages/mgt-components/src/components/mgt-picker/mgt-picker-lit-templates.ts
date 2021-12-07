import { Person, Team, Channel } from '@microsoft/microsoft-graph-types';
import { TemplateResult, html } from 'lit-html';
import { repeat } from 'lit-html/directives/repeat';
import { personMenuContentTemplate, channelMenuTemplate, itemContentsTemplate } from './mgt-picker-fast-templates';
import { MgtPicker, SelectedEntity } from './mgt-picker';
import { DropdownItem } from '../../graph/graph.teams-channels';

function personPickerMenuOption(person: Person, picker: MgtPicker): TemplateResult {
  return html`
    <fast-picker-menu-option
        .contentsTemplate="${personMenuContentTemplate(person)}"
        value="person-${person.id}"
        personThing="Randpm"
        @click=${event => picker.handlePickerMenuClick(event, 'people', person)}>
    </fast-picker-menu-option>
    `;
}

function peoplePickerRepeatTemplate(picker: MgtPicker): TemplateResult {
  return html`
      <div class="entity-text">People</div>
      ${repeat(
        picker.people,
        x => x.id,
        val => personPickerMenuOption(val, picker)
      )}
    `;
}

function channelPickerOption(channel: DropdownItem, team: Team, picker: MgtPicker): TemplateResult {
  return html`
    <fast-picker-menu-option
      .contentsTemplate="${channelMenuTemplate}"
      value="channel-${team.displayName}- ${channel.item.displayName}"
      @click=${event => picker.handlePickerMenuClick(event, 'channel', channel)}>
    </fast-picker-menu-option>
  `;
}

function channelPickerMenuOption(item: DropdownItem, picker: MgtPicker): TemplateResult {
  return html`
    ${repeat(
      item.channels,
      x => x,
      val => channelPickerOption(val, item.item, picker)
    )}
    `;
}

function channelPickerRepeatTemplate(picker: MgtPicker): TemplateResult {
  return html`
      <div class="entity-text">Channels</div>
      
      ${repeat(
        picker.teamItems,
        x => x,
        val => channelPickerMenuOption(val, picker)
      )}
    `;
}

export function pickerDropDownMenuTemplate(picker: MgtPicker): TemplateResult {
  return html`
  <style>
    #custom-menu{
      background: #FFFFFF;
      min-height: 250px;
      max-height: 100%;
      width: inherit;
      color: black;
      padding: 0px;
      /* box-shadow: 0px 1.6px 3.6px rgba(0, 0, 0, 0.13), 0px 1.6px 3.6px rgba(0, 0, 0, 0.13); */
      border-radius: 0px 0px 4px 4px;
      margin-top: 7px;
      filter: drop-shadow()(0px 12.8px 28.8px rgba(0, 0, 0, 0.13)) drop-shadow(0px 0px 9.2px rgba(0, 0, 0, 0.11))
    }
    fast-picker {
      width: 368px;
    }

    fast-picker-menu-option{
      background: #FFFFFF;
      font-family: var(--font-family,var(--default-font-family));
      font-size: var(--font-size,14px);
      line-height: var(--line-height,20px);
      color: rgba(0, 0, 0, 0.83);
      position: inherit;
      padding: 0px;
      height: 44px;
      background: #FFFFFF;
      border-radius: 0px;
    }

    fast-picker-menu-option:hover{
      background: rgba(0, 0, 0, 0.05);
      border-radius: 2px;
    }

    .entity-text {
      width: 36px;
      height: 16px;
      font-family: var(--font-family,var(--default-font-family));
      padding: 9px;
      font-size: var(--font-size,12px);
      line-height: var(--line-height,16px);
      color: rgba(0, 0, 0, 0.55);
    }

    fast-picker-list-item{
      color: #000000;
      width: auto;
      min-width: 32px;
      height: 24px;
      background: rgba(0, 0, 0, 0.07);
      border-radius: 20px;
      border: none;
      padding: 0px;
    }

    .channel{
      display: flex;
    }
    .channel-image p{
      border: var(--avatar-border,0);
      border-radius: var(--avatar-border-radius,50%);
      position: relative;
      width: 24px;
      height: 24px;
      background: #767676;
      color: #FFFFFF;
      text-align: center;
      line-height: 24px;
    }
    .channel-name{
      position: inherit;
      width: inherit;
      padding: 15px 0px 12px 4px;
      font-family: var(--font-family,var(--default-font-family));
      font-weight: var(--font-weight,600);
      font-size: var(--font-size,14px);
      line-height: 20px;
      color: rgba(0, 0, 0, 0.83);
    }
    .not-found-text{
      padding: 9px 0px 9px 13px;
      width: 145px;
      height: 16px;
      font-family: Segoe UI;
      font-size: 12px;
      line-height: 16px;
      color: rgba(0, 0, 0, 0.55);
    }
  </style>
      <fast-picker-menu id="custom-menu">
      ${picker.hasPeople ? peoplePickerRepeatTemplate(picker) : picker.renderNoResultsFound('People')}
      ${picker.hasChannels ? channelPickerRepeatTemplate(picker) : picker.renderNoResultsFound('Channel')}
      </fast-picker-menu>
    `;
}

export function renderFastPickerInput(picker: MgtPicker): TemplateResult {
  return html`
    <style>
      fast-picker input {
        position: inherit;
        width: 348px;
        height: 20px;
        left: 0px;
        top: 0px;
        background: #FFFFFF;
        border: none;
        font-family: Segoe UI;
        font-size: 14px;
        line-height: 20px;
        color: #2B2B2B;
        padding: 0px 0px 0px 1px;
        margin-top: 2px;
      }

      fast-picker-list {
        width: inherit;
        position: static;
        width: 352px;
        height: 24px;
        left: calc(50% - 352px/2);
        top: 6px;
      }

      /**TODO: find out why this is not working */
      fast-picker input::focus{
        border: 2px solid #605E5C;
        box-sizing: border-box;
      }
    </style>
    <!-- Allow single-selection only by setting max-selected="1"-->
    <fast-picker
      max-selected="1"
      no-suggestions-text="No suggestions available"
      suggestions-available-text="Suggestions available"
      loading-text="Loading"
      placeholder="Start typing to search people, chats, and channels"
      filter-selected="false"
      filter-query="false"
      @querychange=${picker.queryChanged}
      @keyup="${picker.onUserKeyUp}"
      .showLoading=${picker.isLoading}
      .listItemContentsTemplate=${itemContentsTemplate}>
        ${pickerDropDownMenuTemplate(picker)}
  </fast-picker>
  `;
}

export function renderEntityBox(picker: MgtPicker): TemplateResult {
  return html`
    <style>
      .entity-box {
        position: inherit;
        width: 368px;
        background: #FFFFFF;
        border: 2px solid #605E5C;
        box-sizing: border-box;
        border-radius: 2px;
        padding: 2px 4px 2px 4px;
        flex: display;
        flex-direction: row;
        align-items: flex-start;
      }
    </style>
    <div class="entity-box">
      ${renderFastPickerInput(picker)}
    </div>
  `;
}
