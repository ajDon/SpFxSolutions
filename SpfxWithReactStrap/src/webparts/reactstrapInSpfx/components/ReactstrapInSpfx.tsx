import * as React from 'react';
import * as jQuery from 'jquery';
import * as bootstrap from 'bootstrap';
import styles from './ReactstrapInSpfx.module.scss';
import 'bootstrap/dist/css/bootstrap.min.css';
import Modal from "./Modal";
import { Button } from "reactstrap";
import { IReactstrapInSpfxProps } from './IReactstrapInSpfxProps';
import pnp from "sp-pnp-js";
import { SPHttpClient } from '../../../../node_modules/@microsoft/sp-http';
import PeoplePicker from './PeoplePicker';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';


export default class ReactstrapInSpfx extends React.Component<IReactstrapInSpfxProps, {
  modal: boolean,
  users: IPersonaProps[],
  IsEditForm: boolean,
  loadPeoplePicker: boolean
}> {
  constructor(props) {
    super(props);
    var queryParameters = new UrlQueryParameterCollection(window.location.href);
    let ItemId = queryParameters.getValue("ItemId");
    this.state = {
      modal: false,
      users: [],
      IsEditForm: ItemId !== "" && ItemId !== undefined,
      loadPeoplePicker: false
    }
    this.showModal = this.showModal.bind(this);
    this._loadCurrentUser = this._loadCurrentUser.bind(this);
    this._renderPeoplePicker = this._renderPeoplePicker.bind(this);
  }

  public componentWillMount() {
    this._loadCurrentUser();
  }

  private _loadCurrentUser() {
    pnp.sp.web.currentUser
      .get()
      .then((respUser) => {
        console.info(respUser);
        var users: IPersonaProps[] = [];
        var user: IPersonaProps = {
          id: respUser.Id.toString(),
          primaryText: respUser.Title,
          secondaryText: respUser.Email,
          imageUrl: `/_layouts/15/userphoto.aspx?size=S&username=${respUser.Email}`
        }
        users.push(user);
        pnp.sp.web.ensureUser("swt.jangale09@ajaysahu.onmicrosoft.com")
          .then((respUser) => {
            var user: IPersonaProps = {
              id: respUser.data.Id.toString(),
              primaryText: respUser.data.Title,
              secondaryText: respUser.data.Email,
              imageUrl: `/_layouts/15/userphoto.aspx?size=S&username=${respUser.data.Email}`
            }
            users.push(user);
            this.setState({
              users: [...this.state.users, ...users],
              loadPeoplePicker: true
            });
          })
      })

  }
  public showModal() {
    this.setState({ modal: !this.state.modal });
    var elementId = jQuery(event.target).attr("data-target");
    jQuery(elementId).modal("show");
  }

  private _renderPeoplePicker() {
    if (this.state.IsEditForm) {
      if (this.state.loadPeoplePicker) {
        return (
          <PeoplePicker
            description="test"
            spHttpClient={this.props.spHttpClient}
            siteUrl={this.props.siteUrl}
            principalTypeUser={true}
            principalTypeSharePointGroup={true}
            principalTypeSecurityGroup={false}
            principalTypeDistributionList={false}
            numberOfItems={20}
            defaultSelectedItems={this.state.users}
          />
        );
      }
    }
    else {
      return (
        <PeoplePicker
          description="test"
          spHttpClient={this.props.spHttpClient}
          siteUrl={this.props.siteUrl}
          principalTypeUser={true}
          principalTypeSharePointGroup={true}
          principalTypeSecurityGroup={false}
          principalTypeDistributionList={false}
          numberOfItems={20}
        />
      );
    }
  }


  public render(): React.ReactElement<IReactstrapInSpfxProps> {
    return (
      <div className={styles.reactstrapInSpfx + " Container"}>

        {this._renderPeoplePicker()}
        <Modal
          ref="modal"
          modalBody="Body"
          modalId="myModal"
          modalFooter=""
          modalTitle="Test Title"
          showClose={true}
          showFooter={false}
        >
          <div>
            <Button color="dark">Test</Button>
          </div>
        </Modal>

        <Modal
          ref="modal"
          modalBody="Body"
          modalId="myModal2"
          modalFooter=""
          modalTitle="Test Title"
          showClose={true}
          showFooter={false}
        >
          <div>
            <Button color="dark">Test 2</Button>
          </div>
        </Modal>
        <Button color="dark" onClick={this.showModal} data-toggle="modal" data-target="#myModal">Launch Modal</Button>
        <Button color="dark" onClick={this.showModal} data-toggle="modal" data-target="#myModal2">Launch Modal</Button>

      </div>
    );
  }
}
