import { html, ref, ViewTemplate } from '@microsoft/fast-element';
import { MgtGenericPicker } from './mgt-picker';

const itemContentsTemplate: ViewTemplate = html`
    <mgt-person user-id="${x => x.value}" view="oneLine"></mgt-person>
`;

const optionContentsTemplate: ViewTemplate = html`
    <mgt-person
        user-id="${x => x.value}"
        view="twoLines"
        line2-property="jobTitle"
    ></mgt-person>
`;

/**
 * The template for the People Picker component.
 * @public
 */
export const peoplePickerTemplate: (context, definition) => ViewTemplate<MgtGenericPicker> = (context, definition) => {
  const domTag = context.tagFor(MgtGenericPicker);

  return html<MgtGenericPicker>`
  <template>
    <${domTag}
        ${ref('picker')}
        :listItemContentsTemplate=${itemContentsTemplate}
        :menuOptionContentsTemplate=${optionContentsTemplate}
        :optionsList=${x => x.optionsList}
        :showLoading=${x => x.showLoading}
        max-selected=${x => x.maxSelected}
        selection=${x => x.selection}
        no-suggestions-text=${x => x.noSuggestionsText}
        suggestions-available-text=${x => x.suggestionsAvailableText}
        loading-text=${x => x.loadingText}
        @menuopening="${(x, c) => x.handleMenuOpening(c.event as Event)}"
        @menuclosing="${(x, c) => x.handleMenuClosing(c.event as Event)}"
        @selectionchange="${(x, c) => x.onPickerSelectionChange(c.event as Event)}"
        @querychange="${(x, c) => x.onPickerQueryChange(c.event as Event)}">
    </$<domTag}>
  </template>
  `;
};
