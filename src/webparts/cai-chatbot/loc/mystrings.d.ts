/**
 * Storage of the web part's localized strings.
 *
 * @interface ChatbotStringsStorage
 * @property description - Web part's description displayed on the placeholder.
 * @property buttonConfigure - The placeholder's configuration button's text.
 * @property propertyPaneDescription - Web part's description displayed on the configuration panel.
 * @property propertyPaneSelect - Channel selection option label's text.
 */
declare interface ChatbotStringsStorage {
  description: string;
  buttonConfigure: string;
  propertyPaneDescription: string;
  propertyPaneSelect: string;
}

declare module 'ChatbotStrings' {
  const strings: ChatbotStringsStorage;
  export = strings;
}
