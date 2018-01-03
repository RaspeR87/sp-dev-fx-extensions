import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { SiteBreadcrumb } from "@pnp/spfx-controls-react/lib/SiteBreadcrumb";

export default class HelloWorld extends React.Component<IHelloWorldProps, {}> {
  public render(): React.ReactElement<IHelloWorldProps> {
    return (
        <div className={ styles.helloWorld }>
            <SiteBreadcrumb context={this.props.context} />
        </div>
    );
  }
}
