import {
  SPHttpClient
} from '@microsoft/sp-http';

export interface IReactstrapInSpfxProps {
  description: string;
  spHttpClient: SPHttpClient;
  siteUrl: string;
}
