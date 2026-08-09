// British-English spellings (SPEC-035 AC-10/AC-10.1) — see enUS.ts.
import { BASE_WORDS } from "./enBase";

const GB_ONLY: readonly string[] = (
  "colour colours coloured colouring favourite favour honour honoured " +
  "humour labour neighbour neighbours rumour valour vigour behaviour " +
  "flavour odour armour " +
  "centre centres centred theatre theatres metre metres litre litres " +
  "fibre calibre sombre lustre " +
  "organise organised organising organisation organisations realise " +
  "realised recognise recognised apologise apologised analyse analysed " +
  "criticise criticised emphasise emphasised " +
  "travelling travelled traveller modelling modelled cancelled " +
  "cancelling labelled labelling signalled fuelled " +
  "defence defences offence licence licensed programme programmes " +
  "catalogue dialogue grey mould moulded jewellery pyjamas doughnut " +
  "aeroplane tyre practise cheque plough"
).split(/\s+/);

export const WORDS: readonly string[] = [...BASE_WORDS, ...GB_ONLY];
