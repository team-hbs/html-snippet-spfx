import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IHtmlSnippetSpfxProps {
  description: string;
  listName: string;
  item: string;
  context: WebPartContext;
  rawHtml:string;
  title:string;
}
