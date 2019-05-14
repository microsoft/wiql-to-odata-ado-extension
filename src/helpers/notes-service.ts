/**
 * Copyright (c) Microsoft Corporation.  All rights reserved.
 */

export interface INote {
    level: 'info' | 'warning' | 'error';
    message: string;
}

export class NotesService {
    private static instanceInternal: NotesService;
    private notesList: INote[] = [];

    /**
     * @returns the instance (current or new) of NotesService.
     */
    public static get instance(): NotesService {
        if (this.instanceInternal == null) {
            this.instanceInternal = new NotesService();
        }
        return this.instanceInternal;
    }

    /**
     * @returns gets all deduped notes
     */
    public get notes(): INote[] {
        // Return unique notes.
        return this.notesList.filter((obj, pos, arr) => {
            return arr.map((note) => note.message).indexOf(obj.message) === pos;
        });
    }

    /**
     * Clears all notes
     */
    public clearAllNotes() {
        this.notesList = [];
    }

    /**
     * Adds a new note
     * @param level info and warning shows the query in the UI. If there are any errors, the query is not show in the UI.
     * @param message the message string to show in the UI.
     */
    public newNote(level: 'info' | 'warning' | 'error', message: string) {
        this.notesList.push({ level, message });
    }
}
