import * as React from "react";
import styles from "./AlertReactJs.module.scss";
import { escape } from "@microsoft/sp-lodash-subset";
import {
  MessageBar,
  MessageBarType,
  Shimmer,
  Label,
  IBreadcrumbItem,
  Breadcrumb,
  ITheme,
  IRenderFunction,
  Link,
  TooltipHost,
  TooltipOverflowMode,
  Stack,
  nullRender,
  Checkbox,
  Modal,
  getTheme,
  mergeStyleSets,
  FontWeights,
} from "office-ui-fabric-react";
import { IItemAddResult, sp, Web } from "@pnp/sp/presets/all";

import { IReadonlyTheme } from "@microsoft/sp-component-base";

import { WebPartTitle } from "@pnp/spfx-controls-react";
import { DisplayMode } from "@microsoft/sp-core-library";
import { IAlertReactJsProps } from "./IAlertReactJsProps";

const theme = getTheme();
const contentStyles = mergeStyleSets({
	container: {
	  display: 'flex',
	  flexFlow: 'column nowrap',
	  alignItems: 'stretch',
	},
	header: [
	  // eslint-disable-next-line deprecation/deprecation
	  theme.fonts.xLargePlus,
	  {
		flex: '1 1 auto',
		borderTop: `4px solid ${theme.palette.themePrimary}`,
		color: theme.palette.neutralPrimary,
		display: 'flex',
		alignItems: 'center',
		fontWeight: FontWeights.semibold,
		padding: '12px 12px 14px 24px',
	  },
	],
	body: {
	  flex: '4 4 auto',
	  padding: '0 24px 24px 24px',
	  overflowY: 'hidden',
	  selectors: {
		p: { margin: '14px 0' },
		'p:first-child': { marginTop: 0 },
		'p:last-child': { marginBottom: 0 },
	  },
	},
  });

export default class Welcome extends React.Component<
  IAlertReactJsProps,
  {
    currentUser: any;
    hasreadmessage: boolean;
    loading: boolean;
    isopen: boolean;
  }
> {
  constructor(props: IAlertReactJsProps) {
    super(props);
    this.state = {
      currentUser: null,
      hasreadmessage: false,
      loading: true,
      isopen: true,
    };
    this.fetchdata = this.fetchdata.bind(this);
    this._onChange = this._onChange.bind(this);
  }

  

  private async fetchdata() {
    let _localuser: any = null;
    await sp.web.currentUser.get().then((r: any) => {
      _localuser = r;
    });

    // Get Item from List
    if (this.props.enablereadmessage) {
      const _list = sp.web.lists.getById(this.props.messagelistid)
      let readitems = await _list.items
        .top(1)
        .filter(`Title eq '${_localuser.LoginName}'`)
        .get();
      this.setState({
        hasreadmessage: readitems.length > 0,
        currentUser: _localuser,
        loading: false,
      });
    } else {
      this.setState({ currentUser: _localuser, loading: false });
    }
  }

  public componentDidMount() {
    this.fetchdata();
  }

  private async _onChange(
    ev: React.FormEvent<HTMLElement>,
    isChecked: boolean
  ) {
    // Add a record to the list
    const _list = sp.web.lists.getById(this.props.messagelistid);
    const iar: IItemAddResult = await _list.items.add({
      Title: this.state.currentUser.LoginName,
    });
    this.setState({ hasreadmessage: true });
  }

  public render(): React.ReactElement<IAlertReactJsProps> {
    const { semanticColors }: IReadonlyTheme = this.props.themeVariant;

    if (this.state.loading) {
      return (
        <div>
          <Shimmer />
        </div>
      );
    } else {
      let lShow = true;

      if (this.props.enablestartdate) {
        if (new Date() < this.props.startdate.value) {
          lShow = false;
        }
      }

      if (this.props.enableenddate) {
        if (new Date() > this.props.enddate.value) {
          lShow = false;
        }
      }

      // Here we need to search the database and get a record back
      if (this.props.enablereadmessage) {
        // Get Current User from Context
        // Look up in the list for a matching record
        if (this.state.hasreadmessage) {
          lShow = false;
        }
      }

      if (!lShow) {
        return null;
      }

      let content = 
        <div style={{padding: "24px"}}>
          {this.props.showtitle ||
          this.props.displayMode === DisplayMode.Edit ? (
            <WebPartTitle
              displayMode={this.props.displayMode}
              title={this.props.title}
              updateProperty={this.props.updateProperty}
              themeVariant={this.props.themeVariant}
            />
          ) : null}
          {this.props.showbodyashtml ? (
            <div dangerouslySetInnerHTML={{ __html: this.props.body }} />
          ) : (
            this.props.body
          )}
          {this.props.enablereadmessage && (
            <Checkbox
              onChange={this._onChange}
              label={this.props.readmessage}
            ></Checkbox>
          )}
        </div>;
      

      let oBody = this.props.showasmodal ? (
        <Modal
          isOpen={this.state.isopen}
          isBlocking={false}
          onDismiss={() => {
            this.setState({ isopen: false });
		  }}
        >{content}</Modal>
      ) : (
        content
	  );
	  
	  return oBody;
    }
  }
}
