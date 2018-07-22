import * as React from 'react';
import * as jQuery from 'jquery';
import * as bootstrap from 'bootstrap';
import styles from './ReactstrapInSpfx.module.scss';
import 'bootstrap/dist/css/bootstrap.min.css';
import Modal from "./Modal";
import { Button } from "reactstrap";
import { IReactstrapInSpfxProps } from './IReactstrapInSpfxProps';
import { SPHttpClient } from '../../../../node_modules/@microsoft/sp-http';
import PeoplePicker from './PeoplePicker';


export default class ReactstrapInSpfx extends React.Component<IReactstrapInSpfxProps, {
  modal: boolean
}> {
  constructor(props) {
    super(props);
    this.state = {
      modal: false
    }
    this.showModal = this.showModal.bind(this);
  }
  public showModal() {
    this.setState({ modal: !this.state.modal });
    var elementId = jQuery(event.target).attr("data-target");
    jQuery(elementId).modal("show");
  }


  public render(): React.ReactElement<IReactstrapInSpfxProps> {
    return (
      <div className={styles.reactstrapInSpfx + " Container"}>
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
