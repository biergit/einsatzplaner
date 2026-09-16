/// <reference path="ConfigTypes.ts" />

type MailJobType = 'einsatz' | 'export';

interface MailJob {
  type: MailJobType;
  /** E-Mail des auslösenden Benutzers (Empfänger des Exports) */
  requester: string;
}

const PENDING_MAIL_JOBS = 'PENDING_MAIL_JOBS';
const MAIL_POLLER_FUNCTION = 'processMailJobs';
const MAIL_POLLER_OWNER = 'MAIL_POLLER_OWNER';

/**
 * Legt einen Mail-Job in die Warteschlange. Der Versand erfolgt durch den
 * Host-Poller (processMailJobs), damit alle Mails vom Host-Konto kommen.
 * Ein bereits wartender Job gleichen Typs wird ersetzt.
 */
function enqueueMailJob(type: MailJobType): void {
  const props = PropertiesService.getScriptProperties();
  const jobs: MailJob[] = JSON.parse(props.getProperty(PENDING_MAIL_JOBS) || '[]');
  const rest = jobs.filter(j => j.type !== type);
  rest.push({ type, requester: Session.getActiveUser().getEmail() });
  props.setProperty(PENDING_MAIL_JOBS, JSON.stringify(rest));
}

/**
 * Wird vom Host-Poller (1-Minuten-Trigger) aufgerufen und versendet
 * ausstehende Mails als Host-Konto. Die Warteschlange wird vor dem Senden
 * geleert, damit bei Fehlern keine Duplikate entstehen.
 */
function processMailJobs(): void {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(PENDING_MAIL_JOBS);
  if (!raw) return;
  props.deleteProperty(PENDING_MAIL_JOBS);

  let jobs: MailJob[] = [];
  try {
    jobs = JSON.parse(raw);
  } catch (e) {
    Logger.log(`processMailJobs: ungültige Warteschlange — ${e}`);
    return;
  }

  for (const job of jobs) {
    try {
      if (job.type === 'einsatz') sendEinsatzEmails();
      else if (job.type === 'export') exportAllData(job.requester);
    } catch (e) {
      Logger.log(`processMailJobs: ${job.type} fehlgeschlagen — ${e}`);
    }
  }
}

/** true, wenn der Host-Poller als Trigger eingerichtet ist. */
function isMailPollerActive(): boolean {
  return ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === MAIL_POLLER_FUNCTION);
}

/** Konto, das den Mail-Versand zuletzt eingerichtet hat (für die Anzeige). */
function getMailPollerOwner(): string {
  return PropertiesService.getScriptProperties().getProperty(MAIL_POLLER_OWNER) || '';
}

/**
 * Richtet den Mail-Poller ein (einmalig vom Host auszuführen) und legt den
 * onEdit-Trigger neu an, damit auch Änderungs-Mails vom Host-Konto kommen.
 * Gibt das Konto zurück, das künftig versendet.
 */
function setupMailPoller(): string {
  for (const t of ScriptApp.getProjectTriggers()) {
    if (t.getHandlerFunction() === MAIL_POLLER_FUNCTION) ScriptApp.deleteTrigger(t);
  }
  ScriptApp.newTrigger(MAIL_POLLER_FUNCTION).timeBased().everyMinutes(1).create();
  reinstallOnEditTrigger();

  const email = Session.getActiveUser().getEmail();
  PropertiesService.getScriptProperties().setProperty(MAIL_POLLER_OWNER, email);
  return email;
}
