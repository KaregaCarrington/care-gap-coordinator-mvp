import { norm, dateVal, fmt } from "./reviewUtils";
import type { Row, Rule } from "./reviewUtils";

export function findMatchingPatient(
  payerRow: Row,
  emr: Row[],
  mapping: any,
  nameKey: string,
  dobKey: string,
) {
  const payerName = norm(payerRow[nameKey]);
  const payerDob = dateVal(payerRow[dobKey]);

  return emr.find((e: Row) => {
    const emrName = norm(e[mapping.name]);
    const emrDob = dateVal(e[mapping.dob]);

    return (
      emrName === payerName &&
      (payerDob !== null &&
        emrDob !== null &&
        emrDob.toDateString() === payerDob.toDateString())
    );
  });
}

export function findEvidence(
  emr: Row[],
  patientName: string,
  mapping: any,
  rule: Rule | undefined
) {
  return emr
    .filter((e: Row) => norm(e[mapping.name]) === patientName)
    .map((e: Row) => ({
      e,
      date: dateVal(e[mapping.date]),
      text: `${e[mapping.description] || ''} ${e[mapping.code] || ''} ${e[mapping.result] || ''}`.toLowerCase()
    }))
    .filter(x =>
      x.date && x.date.getTime() <= Date.now() &&
      (!rule ||
        ((Date.now() - x.date.getTime()) <=
          ((rule.lookbackDays || 365) * 86400000))) &&
      (rule
        ? rule.evidenceTerms.some((t: string) =>
            x.text.includes(t.toLowerCase())
          )
        : false)
    )
    .sort((a, b) => b.date!.getTime() - a.date!.getTime());
}

export function evaluateEvidence(
  evidence: any[],
  measure: string,
  rule: Rule | undefined
) {
  const found = evidence[0];

  if (!found) {
    return {
      status: 'Open',
      notes: 'Order needed'
    };
  }

  const note = (rule?.noteTemplate || `${measure} completed {date}`)
    .replace('{date}', fmt(found.date));

  return {
    status: 'Closed',
    notes: note
  };
}