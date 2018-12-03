import * as React from 'react';
import styles from './HtmlSnippetSpfx.module.scss';
import { IHtmlSnippetSpfxProps } from './IHtmlSnippetSpfxProps';
import { escape } from '@microsoft/sp-lodash-subset';
import pnp, { List, Item, ListEnsureResult, ItemAddResult, FieldAddResult, Web } from "sp-pnp-js"


export default class HtmlSnippetSpfx extends React.Component<IHtmlSnippetSpfxProps, {}> {

  /*
  public componentDidMount(): void {
    console.log("componentDidMount()");
    this.getDocumentContent();
  }

  private getDocumentContent() : Promise<void> { 
    return new Promise<void>((resolve: () => void, reject: (err: string) => void): void => {
      console.log(this.props.context.pageContext.web.absoluteUrl);
     
      let web = new Web(this.props.context.pageContext.web.absoluteUrl);
      web.getFileByServerRelativeUrl(this.props.item).getText().then((text: string) => {
          console.log(text);
        
          resolve();
      });


     
    });
  }
  */



  public render(): React.ReactElement<IHtmlSnippetSpfxProps> {

    //list.items.getById(id).then
    console.log("Render()");

//   <span dangerouslySetInnerHTML={{ __html: this.props.rawHtml }}></span>
    
    return (
      <div className={styles.htmlSnippetSpfx}>
        <div className={styles.container}>
          <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>
            
             
              <span dangerouslySetInnerHTML={{ __html: this.props.rawHtml }}></span>
              
              
          
          </div>
        </div>
      </div>
    );
  }



  
}
