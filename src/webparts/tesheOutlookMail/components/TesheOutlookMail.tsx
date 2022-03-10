import * as React from "react";
import styles from "./TesheOutlookMail.module.scss";
import * as strings from "TesheOutlookMailWebPartStrings";
import {
  ITesheOutlookMailProps,
  IMessage,
  IMessages,
  ITesheOutlookMailState,
} from ".";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import {
  Spinner,
  SpinnerSize,
} from "office-ui-fabric-react/lib/components/Spinner";
import { List } from "office-ui-fabric-react/lib/components/List";
import { Link } from "office-ui-fabric-react/lib/components/Link";
import { IIconProps } from "office-ui-fabric-react/lib/components/Icon";
import { ActionButton } from "office-ui-fabric-react/lib/components/Button";
import { ColorClassNames, Persona, PersonaSize } from "office-ui-fabric-react";
import "react-perfect-scrollbar/dist/css/styles.css";
import { Icon } from "office-ui-fabric-react";
import PerfectScrollbar from "react-perfect-scrollbar";
import { confirmAlert } from "react-confirm-alert"; // Import
import "react-confirm-alert/src/react-confirm-alert.css"; // Import css
import * as $ from "jquery";
import "bootstrap/dist/css/bootstrap.css";
import Dropdown from "react-bootstrap/Dropdown";
// import { Scrollbars } from "react-custom-scrollbars";
import { Scrollbars } from "react-custom-scrollbars-2";
import DropdownItem from "react-bootstrap/esm/DropdownItem";
// import * as Popper from "popper.js";
// import "jquery";
// import "@popperjs/core"; // Edit here
// import "bootstrap/dist/js/bootstrap.bundle";
// import "bootstrap/dist/css/bootstrap.min.css";
// import Dropdown from "react-bootstrap/Dropdown";
// import "bootstrap/dist/js/popper.min.js";

// import Dropdown from "react-bootstrap/Dropdown";
export class TesheOutlookMail extends React.Component<
  ITesheOutlookMailProps,
  ITesheOutlookMailState
> {
  protected readonly outlookLink: string = "https://outlook.office.com/owa/";
  protected readonly outlookNewEmailLink: string =
    "https://outlook.office.com/mail/deeplink/compose";
  renderView: any;
  handleScrollStart: any;

  constructor(props: ITesheOutlookMailProps) {
    super(props);

    this.state = {
      messages: [],
      loading: false,
      error: undefined,
    };
  }

  private addIcon: IIconProps = { iconName: "Add" };
  private viewList: IIconProps = { iconName: "AllApps" };

  /**
   * Load recent messages for the current user
   */
  private _loadMessages(): void {
    if (!this.props.graphClient) {
      return;
    }

    // update state to indicate loading and remove any previously loaded
    // messages
    this.setState({
      error: null,
      loading: true,
      messages: [],
    });

    let graphURI: string = "me/messages";

    if (this.props.showInboxOnly) {
      graphURI = "me/mailFolders/Inbox/messages";
    }

    this.props.graphClient
      .api(graphURI)
      .version("v1.0")
      .select("bodyPreview,receivedDateTime,from,isRead,subject,webLink")
      .top(this.props.nrOfMessages || 15)
      .orderby("receivedDateTime desc")
      .get((err: any, res: IMessages): void => {
        if (err) {
          // Something failed calling the MS Graph
          this.setState({
            error: err.message ? err.message : strings.Error,
            loading: false,
          });
          return;
        }

        // Check if a response was retrieved
        if (res && res.value && res.value.length > 0) {
          this.setState({
            messages: res.value,
            loading: false,
          });
        } else {
          // No messages found
          this.setState({
            loading: false,
          });
        }
      });
  }
  private windowRefresh = () => {
    window.location.reload();
  };
  private readMessage = (id, webLink) => {
    this.props.graphClient
      .api("me/messages/" + id)
      .version("v1.0")
      .update({
        isRead: true,
      })
      .then(() => {
        window.open(webLink, "_blank");
      })
      .then(() => {
        this.windowRefresh();
      });
  };

  private deleteMessage = (id) => {
    this.props.graphClient
      .api("me/messages/" + id)
      .version("v1.0")
      .delete();
    let h = this.state.messages.filter((i) => {
      return i.id !== id;
    });

    this.setState({ messages: h });
  };

  private deletePopup = (id) => {
    confirmAlert({
      title: "Confirm to Delete",
      message: "Are you sure to Delete.",
      buttons: [
        {
          label: "Yes",
          onClick: () => this.deleteMessage(id),
        },
        {
          label: "No",
          onClick: () => console.log("not deleted"),
        },
      ],
    });
  };

  private replyMessage = (id, webLink) => {
    this.props.graphClient
      .api("me/messages/" + id + "/createReply")
      .version("v1.0")
      .post({})
      .then((data) => {
        window.open(data.webLink);
      });
  };
  private replyAllMessage = (id, webLink) => {
    this.props.graphClient
      .api("me/messages/" + id + "/createReplyAll")
      .version("v1.0")
      .post({})
      .then((data) => {
        window.open(data.webLink);
      });
  };

  private forwardMessage = (id, webLink) => {
    this.props.graphClient
      .api("me/messages/" + id + "/createForward")
      .post({})
      .then((data) => {
        window.open(data.webLink);
      });
  };
  private refreshPage = () => {
    window.location.reload();
  };

  /**
   * Render message item
   */

  public openEmail(webLink) {
    window.open(webLink, "_blank");
  }
  public openEmailReply(webLink) {
    window.open(webLink, "_blank");
  }

  private _onRenderCell = (
    item: IMessage,
    index: number | undefined
  ): JSX.Element => {
    // if (item.isRead) {
    //   styles.message = styles.message + " " + styles.isRead;
    // }
    // styles.message = styles.message + " " + styles.isRead;
    type CustomToggleProps = {
      children?: React.ReactNode;
      onClick?: (event: React.MouseEvent<HTMLAnchorElement, MouseEvent>) => {};
    };

    // The forwardRef is important!!
    // Dropdown needs access to the DOM node in order to position the Menu

    const CustomToggle = React.forwardRef(
      (props: CustomToggleProps, ref: React.Ref<HTMLAnchorElement>) => (
        <a
          href=""
          ref={ref}
          onClick={(e) => {
            e.preventDefault();
            props.onClick(e);
          }}
        >
          {props.children}
          <span className={styles.threedots}> </span>
        </a>
      )
    );

    return (
      <div>
        <div className={styles.App}>
          <Dropdown className={styles.dropdown}>
            <Dropdown.Toggle as={CustomToggle} />
            <Dropdown.Menu
              className={styles["dropdownMenu"]}
              size="sm"
              title=""
            >
              <Dropdown.Item
                onClick={() => this.readMessage(item.id, item.webLink)}
              >
                Open
              </Dropdown.Item>
              <DropdownItem
                onClick={() => {
                  this.replyMessage(item.id, item.webLink);
                }}
              >
                Reply
              </DropdownItem>
              <DropdownItem
                onClick={() => {
                  this.replyAllMessage(item.id, item.webLink);
                }}
              >
                Reply All
              </DropdownItem>
              <DropdownItem
                onClick={() => {
                  this.forwardMessage(item.id, item.webLink);
                }}
              >
                Forward
              </DropdownItem>
              <Dropdown.Item
                onClick={() => {
                  this.deletePopup(item.id);
                }}
              >
                Delete
              </Dropdown.Item>
            </Dropdown.Menu>
          </Dropdown>
        </div>

        <Link
          // onClick={() => {
          //   this.readMessage(item.id, item.webLink);
          // }}
          className={
            item.isRead ? `${styles.message} ${styles.isRead}` : styles.message
          }
        >
          <div className={styles.date}>
            {new Date(item.receivedDateTime)
              .toLocaleDateString("en-us", {
                day: "2-digit",
                month: "short",
                hour: "2-digit",
                minute: "2-digit",
                hour12: false,
              })
              .replace(",", "")}
          </div>
          <div className={styles.from}>
            <Persona text={item.from.emailAddress.name} className={styles.from}>
              <a
                href="#"
                onClick={() => {
                  this.readMessage(item.id, item.webLink);
                }}
              >
                <div className={styles.subjects}>{item.subject}</div>
              </a>

              <div className={styles.preview}>{item.bodyPreview}</div>
            </Persona>
          </div>
        </Link>
        
      </div>
    );
  };

  private reocurringCalls: number;
  public componentDidMount(): void {
    // load data initially after the component has been instantiated
    this.reocurringCalls = setInterval(() => {
      this._loadMessages();
    }, 10000);
  }

  public componentDidUpdate(
    prevProps: ITesheOutlookMailProps,
    prevState: ITesheOutlookMailState
  ): void {
    // verify if the component should update. Helps avoid unnecessary re-renders
    // when the parent has changed but this component hasn't
    if (
      prevProps.nrOfMessages !== this.props.nrOfMessages ||
      prevProps.showInboxOnly !== this.props.showInboxOnly
    ) {
      this._loadMessages();
    }
  }

  public render(): React.ReactElement<ITesheOutlookMailProps> {
    const varientStyles = {
      "--varientBGColor": this.props.themeVariant.semanticColors.bodyBackground,
      "--varientFontColor": this.props.themeVariant.semanticColors.bodyText,
      "--varientBGHovered":
        this.props.themeVariant.semanticColors.listItemBackgroundHovered,
    } as React.CSSProperties;

    return (
      <div className={styles.personalEmail} style={varientStyles}>
        <div>Teshe</div>
        <Icon iconName={this.props.iconPicker} />

        <div className={styles.title}>
          <span>
            <Icon className={styles.myIcon} iconName={this.props.iconPicker} />
          </span>
          <h1 className={styles.titleName}>MY OUTLOOK EMAILS</h1>
        </div>
        <Scrollbars
          className={styles.scrollTesh}
          style={{
            width: "100%",
            height: "100%",
            overflow: "hidden",
            // backgroundColor: "green",
          }}
          // thumbMinSize={15}
          // renderView={this.renderView}
          // onScrollStart={this.handleScrollStart}
          // renderThumbVertical={(props) => (
          //   <div {...props} className="thumb-vertical" />
          // )}
          // renderTrackVertical={(props) => (
          //   <div {...props} className="track-vertical" />
          // )}
        >
          <ActionButton
            text={strings.NewEmail}
            iconProps={this.addIcon}
            onClick={this.openNewEmail}
          />
          <ActionButton
            text={strings.ViewAll}
            iconProps={this.viewList}
            onClick={this.openList}
          />
          {this.state.loading && (
            <Spinner label={strings.Loading} size={SpinnerSize.large} />
          )}
          {this.state.messages && this.state.messages.length > 0 ? (
            <div>
              <List
                items={this.state.messages}
                onRenderCell={this._onRenderCell}
                className={styles.list}
              />
            </div>
          ) : (
            !this.state.loading &&
            (this.state.error ? (
              <span className={styles.error}>{this.state.error}</span>
            ) : (
              <span className={styles.noMessages}>{strings.NoMessages}</span>
            ))
          )}
        </Scrollbars>
      </div>
    );
  }

  private openNewEmail = () => {
    window.open(this.outlookNewEmailLink, "_blank");
  };

  private openList = () => {
    window.open(this.outlookLink, "_blank");
  };
}
