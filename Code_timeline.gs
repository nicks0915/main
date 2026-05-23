// Timeline tab backend — fetches the same JIRA data as Assessments
// (same filter 38217, same fields, same pagination logic)
function fetchTimelineDataAsCsv() {
  return fetchAssessmentsDataAsCsv();
}
