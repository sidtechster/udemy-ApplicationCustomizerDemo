import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import { Dialog } from '@microsoft/sp-dialog';

import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from 'AcDemoApplicationCustomizerStrings';

import styles from './ACDemo.module.scss';

const LOG_SOURCE: string = 'AcDemoApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAcDemoApplicationCustomizerProperties {
  // This is an example; replace with your own property
  Top: string;
  Bottom: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AcDemoApplicationCustomizer
  extends BaseApplicationCustomizer<IAcDemoApplicationCustomizerProperties> {

    private _topPlaceHolder: PlaceholderContent | undefined;
    private _bottomPlaceHolder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

    this._renderPlaceHolders();

    return Promise.resolve();
  }

  private _renderPlaceHolders(): void {
    console.log('Availabe placeholders are : ',
      this.context.placeholderProvider.placeholderNames.map(placeholdername => PlaceholderName[placeholdername]).join(', '));

    if (!this._topPlaceHolder) {
      this._topPlaceHolder = 
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top,
          { onDispose: this._onDispose });

      if(!this._topPlaceHolder) {
        console.error('The placeholder Top was not found...');
        return;
      }
      
      if(this.properties) {
        let topString: string = this.properties.Top;
        if (!topString) {
          console.error('Top property was not defined...');
        }

        if(this._topPlaceHolder.domElement) {
          this._topPlaceHolder.domElement.innerHTML = `
            <div class="${styles.acdemoapp}">
              <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.topPlaceHolder}">
                <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> ${escape(topString)}
              </div>
            </div>`;
        }
      }
    }
      
    if (!this._bottomPlaceHolder) {
      this._bottomPlaceHolder = 
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Bottom,
          { onDispose: this._onDispose });

      if(!this._bottomPlaceHolder) {
        console.error('The placeholder Top was not found...');
        return;
      }
      
      if(this.properties) {
        let bottomString: string = this.properties.Bottom;
        if (!bottomString) {
          console.error('Bottom property was not defined...');
        }

        if(this._bottomPlaceHolder.domElement) {
          this._bottomPlaceHolder.domElement.innerHTML = `
            <div class="${styles.acdemoapp}">
              <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.bottomPlaceHolder}">
                <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> ${escape(bottomString)}
              </div>
            </div>`;
        }
      }
    }
    
  }

  private _onDispose(): void {
    console.log('Disposed custom top and bottom placeholders.')
  }
}
