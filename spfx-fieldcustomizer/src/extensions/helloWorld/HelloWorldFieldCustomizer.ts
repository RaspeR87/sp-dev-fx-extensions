import * as jQuery from 'jquery';
import 'jqueryui';

import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import {
  CellFormatter,
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';

import * as strings from 'helloWorldStrings';
import styles from './HelloWorld.module.scss';

import pnp, { List, ItemUpdateResult, Item } from 'sp-pnp-js';
import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IHelloWorldProperties {
}

const LOG_SOURCE: string = 'HelloWorldFieldCustomizer';

export default class HelloWorldFieldCustomizer
  extends BaseFieldCustomizer<IHelloWorldProperties> {

  @override
  public onInit(): Promise<void> {
    SPComponentLoader.loadCss('//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css');

    return Promise.resolve<void>();
  }

  @override
  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    pnp.sp.web.lists.getByTitle(this.context.pageContext.list.title).items.getById(parseInt(event.row.getValueByName('ID').toString()))
        .get()
        .then(r => {
          event.cellDiv.innerHTML = '<div class="spfx-tooltip"><a href="#" title="' + r.AdditionalDescription + '">' + CellFormatter.renderAsText(this.context.column, event.cellValue) + '</a></div>';
        
          jQuery(function() {
            jQuery('.spfx-tooltip').tooltip();
          });
        });
  }

  @override
  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    super.onDisposeCell(event);
  }
}
