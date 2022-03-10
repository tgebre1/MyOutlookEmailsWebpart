export interface IMessages {
  id: any;
  value: IMessage[];
}

export interface IMessage {
  bodyPreview: string;
  from: {
    emailAddress: {
      address: string;
      name: string;
    };
  };
  isRead: boolean;
  receivedDateTime: string;
  subject: string;
  webLink: string;
  id: string;
}
