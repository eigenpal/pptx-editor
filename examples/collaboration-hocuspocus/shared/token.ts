// The demo's Hocuspocus token, which the app sends and the server reads. A real app sends a
// signed token from its sign-in service instead.

export interface RoomToken {
  /** The shared secret. */
  readonly secret: string;
  /** This tab starts the room, so the server may create it. */
  readonly createRoom?: boolean;
}

export function encodeRoomToken(token: RoomToken): string {
  return JSON.stringify(token);
}

/** The token's claims, or null when it is not a token of this demo. */
export function readRoomToken(raw: string): RoomToken | null {
  try {
    const value: unknown = JSON.parse(raw);
    if (typeof value !== 'object' || value === null) return null;
    const { secret, createRoom } = value as Record<string, unknown>;
    if (typeof secret !== 'string') return null;
    return { secret, createRoom: createRoom === true };
  } catch {
    return null;
  }
}
