/**
 * Adds residents to the "Your name" dropdown on the LIVE give-away form
 * "PA CRC — I cannot work a scheduled shift" (file id below).
 * Keeps the existing PA names first, then residents sorted by last name.
 * Safe to re-run: names already in the list are not duplicated.
 * Run once from a standalone Apps Script project (script.new) → approve Forms access.
 */
const GIVEAWAY_FORM_ID = '1Bgl7CygReQqzt3B6ez8Y_m1tM0WO_dykvEvzxdXq2-E';

const RESIDENTS = [
  'Pietro Cianflone',
  'Stephen Dueck',
  'Martina Erdstein',
  'Sophie Gregoire-Mitha',
  'Delanee Hawkins',
  'Bushra Khalid',
  'Amit Manocha',
  'Jules Perez',
  'Will Siemens',
  'Caitlin Wachal'
];

function addResidentsToGiveawayForm() {
  const form = FormApp.openById(GIVEAWAY_FORM_ID);
  const item = form.getItems().find(function (it) {
    return it.getTitle().trim() === 'Your name';
  });
  if (!item) throw new Error('No "Your name" question found on the form.');

  let q;
  const type = item.getType();
  if (type === FormApp.ItemType.LIST) q = item.asListItem();
  else if (type === FormApp.ItemType.MULTIPLE_CHOICE) q = item.asMultipleChoiceItem();
  else throw new Error('"Your name" is type ' + type + ', not a dropdown.');

  const existing = q.getChoices().map(function (c) { return c.getValue(); });
  const have = existing.map(function (n) { return n.toLowerCase(); });
  const toAdd = RESIDENTS.filter(function (n) { return have.indexOf(n.toLowerCase()) === -1; });

  const all = existing.concat(toAdd);
  q.setChoices(all.map(function (n) { return q.createChoice(n); }));

  Logger.log('Added ' + toAdd.length + ': ' + toAdd.join(', '));
  Logger.log('Dropdown now (' + all.length + '): ' + all.join(' | '));
}
