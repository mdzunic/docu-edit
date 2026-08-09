// American-English spellings (SPEC-035 AC-10/AC-10.1's "the dictionary
// matches the text's language" — these forms are valid in en-US and flagged
// in en-GB, and vice versa via enGB.ts's US_SPELLED_DIFFERENTLY list).
import { BASE_WORDS } from "./enBase";

const US_ONLY: readonly string[] = (
  "color colors colored coloring favorite favor honor honored humor labor " +
  "neighbor neighbors rumor valor vigor behavior flavor odor armor " +
  "center centers centered theater theaters meter meters liter liters " +
  "fiber caliber somber luster " +
  "organize organized organizing organization organizations realize " +
  "realized recognize recognized apologize apologized analyze analyzed " +
  "criticize criticized emphasize emphasized " +
  "traveling traveled traveler modeling modeled canceled canceling " +
  "labeled labeling signaled fueled " +
  "defense defenses offense license licensed program programs catalog " +
  "dialog gray mold molded jewelry pajamas donut " +
  "airplane tire practice check plow"
).split(/\s+/);

export const WORDS: readonly string[] = [...BASE_WORDS, ...US_ONLY];
