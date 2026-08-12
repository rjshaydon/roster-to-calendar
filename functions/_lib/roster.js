// The browser/manual parser is the canonical roster interpretation. Server
// endpoints import it through this adapter so imports and manual uploads can
// never drift into separately maintained rule sets.
export * from "../../public/static/roster.js";
