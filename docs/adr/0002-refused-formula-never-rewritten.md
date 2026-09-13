# A refused formula is never rewritten by reference

When the parser refuses formula text, the references in it are unknown. Sheet rename and sheet delete therefore leave that text exactly as it is, instead of guessing at it with a regex. Two alternatives were rejected:

- **A regex fallback, as the shifter uses.** It is consistent with the shifter, but it guesses at meaning that was never established.
- **Validate everything, then throw before changing anything.** This makes a rename atomic, but it turns one unusual cell into a rename that always fails.

Leaving the text as it is is the only choice that can neither half-apply a rename nor invent meaning. The shifter's regex fallback, narrowed by spec 25, remains the one sanctioned exception, and it must not spread to other paths. Decided in the round-4 architecture review (specs 54 and 55).
