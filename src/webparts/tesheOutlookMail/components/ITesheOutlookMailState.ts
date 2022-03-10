import { IMessage } from ".";

export interface ITesheOutlookMailState {
  error: string;
  loading: boolean;
  messages: IMessage[];
}
