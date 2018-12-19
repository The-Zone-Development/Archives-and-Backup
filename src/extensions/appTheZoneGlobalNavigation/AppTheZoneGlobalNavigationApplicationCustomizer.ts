import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
//import { Dialog } from '@microsoft/sp-dialog';

import styles from './AppCustomizer.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from 'AppTheZoneGlobalNavigationApplicationCustomizerStrings';
const LOG_SOURCE: string = 'AppTheZoneGlobalNavigationApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAppTheZoneGlobalNavigationApplicationCustomizerProperties {
  Top: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AppTheZoneGlobalNavigationApplicationCustomizer
  extends BaseApplicationCustomizer<IAppTheZoneGlobalNavigationApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
    this._renderPlaceHolders();
    return Promise.resolve();
  }

  private _renderPlaceHolders(): void {
    console.log(
        "Available placeholders: ",
        this.context.placeholderProvider.placeholderNames
            .map(name => PlaceholderName[name])
            .join(", ")
    );

    // Handling the top placeholder
    if (!this._topPlaceholder) {
        this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
            PlaceholderName.Top,
            { onDispose: this._onDispose }
        );

        // The extension should not assume that the expected placeholder is available.
        if (!this._topPlaceholder) {
            console.error("The expected placeholder (Top) was not found.");
            return;
        }

        if (this.properties) {
            let topString: string = this.properties.Top;
            if (!topString) {
                topString = "(Top property was not defined.)";
            }

            if (this._topPlaceholder.domElement) {
                this._topPlaceholder.domElement.innerHTML = `
                <div class="${styles.app}">
                    <div class="${styles.top}">
                        <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> ${escape(
                            topString
                        )}
                    </div>
                </div>`;
            }
        }
    }
  }
  private _onDispose(): void {
    console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}
