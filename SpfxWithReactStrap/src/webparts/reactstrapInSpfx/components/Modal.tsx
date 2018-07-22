import * as React from 'react';

export interface IModal {
    modalId: string;
    modalTitle: string;
    showClose: boolean;
    showFooter: boolean;
    modalBody: string;
    modalFooter: string;
}

export default class Modal extends React.Component<IModal, {
    showModal: boolean
}> {
    constructor(props) {
        super(props);
    }

    public render(): React.ReactElement<IModal> {
        return (
            <div className="modal fade" role="dialog" tabIndex={-1} id={this.props.modalId}>
                <div className="modal-dialog" role="document">
                    <div className="modal-content">
                        <div className="modal-header">
                            <h5 className="modal-title" id="exampleModalLabel">{this.props.modalTitle}</h5>
                            {
                                this.props.showClose === true ?
                                    <button type="button" className="close" data-dismiss="modal" aria-label="Close">
                                        <span aria-hidden="true">&times;</span>
                                    </button>
                                    : ""
                            }

                        </div>
                        <div className="modal-body">
                            {this.props.children}
                        </div>
                        {
                            this.props.showFooter === true ?
                                <div className="modal-footer">
                                    <button className='btn btn-default' data-dismiss='modal' aria-label='Close'>Close</button>
                                </div>
                                : ""
                        }

                    </div>
                </div>
            </div>
        );
    }
}