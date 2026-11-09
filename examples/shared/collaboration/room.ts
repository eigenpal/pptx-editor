// The room helpers that both collaboration examples share: the people you can sign in as, a
// per-tab actor id, room ids, and invite links. Framework-free.

/** One person the examples let you sign in as. */
export interface Person {
  /** The user id, the stable half of the actor id. */
  readonly id: string;
  /** The display name. Remote carets and comments show it. */
  readonly name: string;
  /** The presence color, as a hex string. */
  readonly color: string;
}

/** Pick a different person in each browser. */
export const PEOPLE: readonly Person[] = [
  { id: 'avery', name: 'Avery', color: '#2563eb' },
  { id: 'blake', name: 'Blake', color: '#c2410c' },
  { id: 'casey', name: 'Casey', color: '#15803d' },
  { id: 'devon', name: 'Devon', color: '#9333ea' },
];

/** Random lowercase hex, from `getRandomValues` (it works off a secure origin too). */
export function randomHex(bytes: number): string {
  return [...crypto.getRandomValues(new Uint8Array(bytes))]
    .map((byte) => byte.toString(16).padStart(2, '0'))
    .join('');
}

/**
 * The actor id of one editor attachment. It must be unique per tab, not per person: the same
 * person in two tabs is two replicas.
 */
export function actorIdFor(person: Person): string {
  return `${person.id}#${randomHex(8)}`;
}

/** Room ids: 16 to 64 lowercase letters or digits. */
const ROOM_ID = /^[a-z0-9]{16,64}$/;

/** A new room id. */
export function createRoomId(): string {
  return randomHex(16);
}

/** The room that the address bar names, or null when this tab starts a new room. */
export function roomFromUrl(href: string = location.href): string | null {
  const room = new URL(href).searchParams.get('room');
  return room && ROOM_ID.test(room) ? room : null;
}

/** The address of a room, with an optional fragment (never sent to a server). */
export function roomUrl(roomId: string, fragment = '', href: string = location.href): string {
  const url = new URL(href);
  url.search = '';
  url.searchParams.set('room', roomId);
  url.hash = fragment;
  return url.toString();
}

/** A room id short enough to compare across two windows. */
export function shortRoomId(roomId: string): string {
  return `${roomId.slice(0, 6)}…${roomId.slice(-4)}`;
}
