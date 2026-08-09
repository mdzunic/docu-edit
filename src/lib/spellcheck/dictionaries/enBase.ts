// Common English vocabulary shared by en-US and en-GB (SPEC-035) — a
// curated word list, not a full Hunspell dictionary (see dictionary.ts's
// header comment for why). Deliberately excludes words whose US/GB spelling
// differs (see enUS.ts/enGB.ts); everything here is spelled the same way in
// both.
export const BASE_WORDS: readonly string[] = (
  "a able about above across act action add after again against age ago " +
  "agree all almost alone along already also although always am among " +
  "an and another answer any anyone anything appear are area around as " +
  "ask at attention away back bad bag ball base be beautiful because " +
  "become bed been before began begin behind being believe below best " +
  "better between big bird bit black blue boat body book both box boy " +
  "bring brother build building but buy by call came can car card care " +
  "carry case cat cause certain change check child children choose city " +
  "class clear close cold come common company complete consider contain " +
  "continue control could country course cover create cut dark data day " +
  "days dear decide deep did die different difficult direct do does dog " +
  "done door down draw dream drink drive during each early earth east " +
  "easy eat effect eight either else end enough even evening ever every " +
  "example experience explain eye face fact fall family far farm fast " +
  "father fear feel feet few field figure fill final find fine finger " +
  "finish fire first fish five floor fly follow food foot for force " +
  "form found four free friend from front full fun future game garden " +
  "gave get girl give glass go goes going good got government great " +
  "green ground group grow guess had hand hard has have he head hear " +
  "heard heart heavy help her here high him his history hold home hope " +
  "horse hot hour house how however hundred i idea if important in " +
  "increase indeed information inside instead interest into is issue " +
  "it its itself join just keep kept kind knew know known land language " +
  "large last late later laugh learn least leave left leg less let " +
  "letter level life light like line list listen little live long look " +
  "lot love low made main make man many map matter may maybe me mean " +
  "meet member men might mile mind minute miss moment money month more " +
  "morning most mother move much music must my name near need never new " +
  "next night no nor north not note nothing now number of off often oh " +
  "old on once one only open or order other our out outside over own " +
  "page paper part particular pass past pattern people perhaps person " +
  "picture piece place plan plant play please point poor position " +
  "possible power present pretty probably problem produce program " +
  "public pull push put question quick quiet quite rain reach read " +
  "ready real reason receive record red remember report rest result " +
  "return right river road room round run said same saw say school sea " +
  "season second see seem seen sense series serve set seven several " +
  "shall she short should show side simple since sing sister sit six " +
  "size small so some someone something sometimes song soon sound south " +
  "space speak special stand start state stay step still stop story " +
  "street strong student study such sudden sure system table take talk " +
  "teach tell ten than that the their them then there these they thing " +
  "think third this those though thought three through time to today " +
  "together told too took top toward town train tree trouble true try " +
  "turn two under understand until up upon us use used very view voice " +
  "wait walk want war warm was watch water way we week well went were " +
  "west what when where whether which while white who whole whose why " +
  "will wind window wish with within without woman women word work " +
  "world would write year years yes yet you young your"
).split(/\s+/);
