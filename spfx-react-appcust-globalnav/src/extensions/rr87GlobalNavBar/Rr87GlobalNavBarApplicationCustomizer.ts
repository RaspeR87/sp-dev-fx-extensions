import * as React from 'react';
import * as ReactDom from 'react-dom';

import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import { escape } from '@microsoft/sp-lodash-subset'; 

import GlobalNavBar from './components/GlobalNavBar/GlobalNavBar';
import { IGlobalNavBarProps } from './components/GlobalNavBar/IGlobalNavBarProps';
import GlobalFooterBar from './components/GlobalFooterBar/GlobalFooterBar';
import { IGlobalFooterBarProps } from './components/GlobalFooterBar/IGlobalFooterBarProps';
import * as SPTermStore from '../../components/SPTermStoreService'; 

import styles from './Rr87GlobalNavBarApplicationCustomizer.module.scss';

import * as strings from 'rr87GlobalNavBarStrings';

const LOG_SOURCE: string = 'Rr87GlobalNavBarApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IRr87GlobalNavBarApplicationCustomizerProperties {
  NavTermSet?: string;
  FooterTermSet?: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class Rr87GlobalNavBarApplicationCustomizer
  extends BaseApplicationCustomizer<IRr87GlobalNavBarApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;
  private _bottomPlaceholder: PlaceholderContent | undefined;
  private _topMenuItems: SPTermStore.ISPTermObject[];
  private _bottomMenuItems: SPTermStore.ISPTermObject[];

  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    
    let termStoreService: SPTermStore.SPTermStoreService = new SPTermStore.SPTermStoreService({
      spHttpClient: this.context.spHttpClient,
      siteAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
    });

    if (this.properties.NavTermSet != null) {
      this._topMenuItems = await termStoreService.getTermsFromTermSetAsync(this.properties.NavTermSet);
    }
    if (this.properties.FooterTermSet != null) {
      this._bottomMenuItems = await termStoreService.getTermsFromTermSetAsync(this.properties.FooterTermSet);
    }

    // Call render method for generating the needed html elements
    this._renderPlaceHolders();

    return Promise.resolve<void>();
  }

  private _renderPlaceHolders(): void {
    
    console.log('Available placeholders: ',
      this.context.placeholderProvider.placeholderNames.map(name => PlaceholderName[name]).join(', '));

    // Handling the top placeholder
    if (!this._topPlaceholder) {
      this._topPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top,
          { onDispose: this._onDispose });

      // The extension should not assume that the expected placeholder is available.
      if (!this._topPlaceholder) {
        console.error('The expected placeholder (Top) was not found.');
        return;
      }

      if (this._topMenuItems != null && this._topMenuItems.length > 0) {
        const element: React.ReactElement<IGlobalNavBarProps> = React.createElement(
          GlobalNavBar,
          {
            menuItems: this._topMenuItems,
          }
        );
    
        ReactDom.render(element, this._topPlaceholder.domElement);
      }
    }

    // Handling the bottom placeholder
    if (!this._bottomPlaceholder) {
      this._bottomPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Bottom,
          { onDispose: this._onDispose });

      // The extension should not assume that the expected placeholder is available.
      if (!this._bottomPlaceholder) {
        console.error('The expected placeholder (Bottom) was not found.');
        return;
      }

      if (this._bottomMenuItems != null && this._bottomMenuItems.length > 0) {
        const element: React.ReactElement<IGlobalNavBarProps> = React.createElement(
          GlobalFooterBar,
          {
            menuItems: this._bottomMenuItems,
          }
        );
    
        ReactDom.render(element, this._bottomPlaceholder.domElement);
      }
    }
  }

  private _onDispose(): void {
    console.log('[TenantGlobalNavBarApplicationCustomizer._onDispose] Disposed custom nav and bottom placeholders.');
  }
}
