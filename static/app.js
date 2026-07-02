// Global save helper used by multiple pages
function saveField(id, field, value) {
  return fetch(`/customer/${id}/update`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ [field]: value })
  });
}
