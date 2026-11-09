// The React parts that both collaboration examples share: the seat picker, a full-page
// message, and the editor with its room bar. The examples differ only in transport: each one
// builds a `YjsCollaborationSession` over its own provider and passes it here.

import { useCallback, useEffect, useMemo, useState, useSyncExternalStore } from 'react';
import type { ReactNode } from 'react';
import type { CollaborationParticipant, EditorModule } from '@pptx-editor.dev/core/editor';
import { collaborationModule, type YjsCollaborationSession } from '@pptx-editor.dev/pro';
import { drawingModule } from '@pptx-editor.dev/pro/drawing';
import { PptxEditor, type UsePptxSourceResult } from '@pptx-editor.dev/react';
import { PEOPLE, shortRoomId, type Person } from './room.ts';

/** Pick a person. In a second browser, pick somebody else. */
export function SeatPicker({
  title,
  lede,
  onPick,
}: {
  readonly title: string;
  readonly lede: string;
  readonly onPick: (person: Person) => void;
}) {
  return (
    <main className="room-stage">
      <h1>{title}</h1>
      <p>{lede}</p>
      <ul className="room-seats">
        {PEOPLE.map((person) => (
          <li key={person.id}>
            <button type="button" className="room-seat" onClick={() => onPick(person)}>
              <Avatar name={person.name} color={person.color} />
              {person.name}
            </button>
          </li>
        ))}
      </ul>
    </main>
  );
}

/** A full-page message: a wait, or a failure with what to do next. */
export function RoomMessage({
  title,
  children,
  alert = false,
}: {
  readonly title: string;
  readonly children?: ReactNode;
  readonly alert?: boolean;
}) {
  return (
    <main className="room-stage" role={alert ? 'alert' : 'status'}>
      <h1>{title}</h1>
      {children}
    </main>
  );
}

function Avatar({ name, color }: { readonly name: string; readonly color?: string }) {
  return (
    <span className="room-avatar" style={{ background: color }} aria-hidden="true">
      {[...name][0]?.toUpperCase()}
    </span>
  );
}

/** The session's participants, for React. */
function useParticipants(session: YjsCollaborationSession): readonly CollaborationParticipant[] {
  return useSyncExternalStore(
    useCallback((notify) => session.subscribeParticipants(notify), [session]),
    () => session.participants()
  );
}

/** The session's status snapshot, for React. */
function useStatus(session: YjsCollaborationSession) {
  return useSyncExternalStore(
    useCallback((notify) => session.subscribeStatus(notify), [session]),
    () => session.statusSnapshot()
  );
}

/** What the status says, as a word and a glyph, so it never depends on color alone. */
function statusLabel(status: string): { readonly label: string; readonly tone: string } {
  if (status === 'ready') return { label: '● Live', tone: 'is-live' };
  if (status === 'disconnected') return { label: '○ Offline: edits merge on reconnect', tone: '' };
  if (status === 'error') return { label: '▲ Out of sync', tone: 'is-bad' };
  return { label: '○ Connecting', tone: '' };
}

/**
 * The editor in a room. `session` is a `YjsCollaborationSession` that the example created
 * after its provider loaded the room. The session seeds an empty room from `source`, and
 * replaces `source` with the room's content when the room has some.
 */
export function RoomEditor({
  session,
  source,
  roomId,
  inviteUrl,
  transport,
}: {
  readonly session: YjsCollaborationSession;
  readonly source: UsePptxSourceResult;
  readonly roomId: string;
  /** The link a second person opens to join. */
  readonly inviteUrl: string;
  /** What carries the room, for the bar. */
  readonly transport: string;
}) {
  // Read once at mount: one module list for the editor's life, with this session in it.
  const modules = useMemo<readonly EditorModule[]>(
    () => [collaborationModule({ session }), drawingModule()],
    [session]
  );
  const participants = useParticipants(session);
  const { status, reason } = useStatus(session);
  const state = statusLabel(status);
  const [copied, setCopied] = useState<'yes' | 'failed' | null>(null);
  useEffect(() => {
    if (copied !== 'yes') return undefined;
    const timer = setTimeout(() => setCopied(null), 1600);
    return () => clearTimeout(timer);
  }, [copied]);

  const copyInvite = () => {
    if (!navigator.clipboard) return setCopied('failed');
    navigator.clipboard.writeText(inviteUrl).then(
      () => setCopied('yes'),
      () => setCopied('failed')
    );
  };

  return (
    <div className="room-shell">
      <header className="room-bar">
        <span className={`room-status ${state.tone}`} role="status" title={reason?.code}>
          {state.label}
        </span>
        <span className="room-transport">{transport}</span>
        <button type="button" className="room-invite" onClick={copyInvite} title={inviteUrl}>
          Room <b>{shortRoomId(roomId)}</b>
          <span>
            {copied === 'yes'
              ? 'Link copied'
              : copied === 'failed'
                ? 'Copy the link from the address bar'
                : 'Copy invite link'}
          </span>
        </button>
        <span className="room-spacer" />
        <ul className="room-people" aria-label="People in this room">
          {participants.map((p) => (
            <li key={p.actorId} title={p.isLocal ? `${p.name} (you)` : p.name}>
              <Avatar name={p.name} color={p.color} />
              <span className="room-people__name">{p.isLocal ? `${p.name} (you)` : p.name}</span>
            </li>
          ))}
        </ul>
      </header>
      {status === 'error' && reason?.code === 'schema-version-mismatch' ? (
        <p className="room-alert" role="alert">
          This room uses another collaboration format. Save your copy, then update every participant
          to the same release.
        </p>
      ) : null}
      <div className="room-editor">
        <PptxEditor
          document={source.document}
          {...(source.fonts ? { fonts: source.fonts } : {})}
          modules={modules}
          author={session.identity.name}
          colorMode="system"
        />
      </div>
    </div>
  );
}
