/**
 * Add "Tanya Sala" to the Name dropdown on both PVCS psychiatrist
 * on-call schedule request forms.
 *
 *   Current period (Aug 31 – Nov 8 2026): 1sHtml5D_u7kYloSqX6W5LAIosT5jhiYqrCImAqskX2w
 *   Next period    (Nov 9 – Jan 3 2027):  1bvojJ3zQLJTN_1nUxp2vVK3KL6gQuU9aysSG52jtoU8
 *
 * Alphabetical by last name — inserts between Ruzhynsky and Sareen.
 * Idempotent: skips a form that already has the name.
 *
 * Paste this file into any Apps Script project owned by the form owner,
 * then Run the addTanyaSala function. First run will ask for FormApp
 * authorization.
 */
function addTanyaSala() {
  var FORM_IDS = [
    '1sHtml5D_u7kYloSqX6W5LAIosT5jhiYqrCImAqskX2w',
    '1bvojJ3zQLJTN_1nUxp2vVK3KL6gQuU9aysSG52jtoU8'
  ];
  var NEW_NAME = 'Tanya Sala';
  var lastNameKey = function (fullName) {
    var parts = fullName.trim().split(/\s+/);
    return parts[parts.length - 1].toLowerCase();
  };
  var log = [];
  FORM_IDS.forEach(function (id) {
    var form = FormApp.openById(id);
    var items = form.getItems(FormApp.ItemType.LIST);
    var nameItem = null;
    for (var i = 0; i < items.length; i++) {
      if (/^\s*Name\s*$/i.test(items[i].getTitle())) {
        nameItem = items[i].asListItem();
        break;
      }
    }
    if (!nameItem) { log.push(id + ': ERROR — no "Name" dropdown found'); return; }
    var current = nameItem.getChoices().map(function (c) { return c.getValue(); });
    if (current.indexOf(NEW_NAME) !== -1) {
      log.push(id + ': already present (' + current.length + ' names)');
      return;
    }
    var updated = current.concat([NEW_NAME]).sort(function (a, b) {
      var la = lastNameKey(a), lb = lastNameKey(b);
      return la < lb ? -1 : la > lb ? 1 : 0;
    });
    nameItem.setChoiceValues(updated);
    log.push(id + ': inserted at ' + (updated.indexOf(NEW_NAME) + 1) + '/' + updated.length);
  });
  Logger.log(log.join('\n'));
  return log;
}
