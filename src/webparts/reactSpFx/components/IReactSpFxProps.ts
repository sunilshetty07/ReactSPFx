import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IReactSpFxProps {
  userDisplayName: string;
  context: WebPartContext;
  selectedList: any;
}
