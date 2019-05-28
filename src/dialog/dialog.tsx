/**
 * Copyright (c) Microsoft Corporation.  All rights reserved.
 */

import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { initializeIcons } from '@uifabric/icons';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { ITextField, TextField } from 'office-ui-fabric-react/lib/TextField';

import { INote, NotesService } from '../helpers/notes-service';
import { IActionContext } from '../models';
import { DialogHelper } from './dialog-helper';

import './dialog.scss';

export interface IAppComponentState {
    notes: INote[];
    oDataString: string;
    loading: boolean;
    copied: boolean;
}

class AppComponent extends React.Component<{}, IAppComponentState> {
    public textField: ITextField;
    public queryName: string;

    public constructor(props: any) {
        super(props);

        initializeIcons();
        this.state = {
            notes: [],
            oDataString: null,
            loading: true,
            copied: false,
        };
    }

    public async componentDidMount(): Promise<void> {
        const config: IActionContext = VSS.getConfiguration();
        this.queryName = config.query.name;
        const dialogHelper = new DialogHelper(config);
        const oDataUrl = await dialogHelper.getODataUrl();
        this.setState({
            notes: NotesService.instance.notes,
            oDataString: oDataUrl,
            loading: false,
        });
    }

    public render() {
        if (this.state.loading) {
            return <Spinner size={SpinnerSize.large} />;
        }

        return (
            <div>
                <h2 className='query-name-heading'>{this.queryName}</h2>
                {!this.state.notes.map((n) => n.level).includes('error') &&
                    <div className='query-area'>
                        <p>Caution: this is not guaranteed to work. Please heed any warnings below and test before use!</p>
                        <TextField
                            id='copyField'
                            multiline rows={8}
                            spellCheck={false}
                            defaultValue={this.state.oDataString}
                            componentRef={(textField) => { this.textField = textField; }}
                        />
                        <div className='button-area'>
                            <PrimaryButton text={this.state.copied ? 'Copied.' : 'Copy query'}
                                className='button'
                                iconProps={{ iconName: this.state.copied ? 'CheckMark' : 'Copy' }}
                                onClick={() => this.copyClicked()}
                            />
                            <PrimaryButton text='Open in new tab'
                                className='button'
                                iconProps={{ iconName: 'OpenInNewWindow' }}
                                onClick={() => this.openClicked()}
                            />
                        </div>
                    </div>}
                <div className='note-area'>
                    {this.state.notes.map((note, index) => {
                        return <MessageBar
                            key={index}
                            className='note-message-bar'
                            messageBarType={MessageBarType[note.level]}
                            isMultiline={false}
                            truncated={true}
                            overflowButtonAriaLabel='See more'
                        >
                            {note.message}
                        </MessageBar>;
                    })}
                </div>
            </div>
        );
    }

    private copyClicked() {
        this.textField.select();
        document.execCommand('copy');
        this.setState({ copied: true });
        setTimeout(() => this.setState({ copied: false }), 3000);
    }

    private openClicked() {
        window.open(this.state.oDataString, '_blank');
    }
}

ReactDOM.render(
    <AppComponent />,
    document.getElementById('react-mount')
);
