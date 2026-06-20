var SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyQxJ56SkHlxRHIcE5Il4bjxzPv6TubQVxeJ8-9VlZeJHTJOOcvVI0VUAWWAzWksIqTlA/exec';

window.addEventListener('DOMContentLoaded', function() {
  document.getElementById('f-date').addEventListener('change', function() {
    var days = ['អាទិត្យ','ចន្ទ','អង្គារ','ពុធ','ព្រហស្បតិ៍','សុក្រ','សៅរ៍'];
    var d = new Date(this.value);
    document.getElementById('f-day').value = days[d.getDay()];
  });
  loadBookings();
});

function showAlert(type, msg) {
  ['err','ok','warn'].forEach(function(t) {
    var el = document.getElementById('alert-'+t);
    el.classList.remove('show'); el.innerHTML = '';
  });
  var el = document.getElementById('alert-'+type);
  el.innerHTML = msg; el.classList.add('show');
  setTimeout(function(){ el.classList.remove('show'); }, 6000);
}

function submitBooking() {
  var dateVal    = document.getElementById('f-date').value;
  var dayVal     = document.getElementById('f-day').value;
  var timeVal    = document.getElementById('f-time').value;
  var roomVal    = document.getElementById('f-room').value;
  var topicVal   = document.getElementById('f-topic').value.trim();
  var groupVal   = document.getElementById('f-group').value.trim();
  var phoneVal   = document.getElementById('f-phone').value.trim();
  var emailVal   = document.getElementById('f-email').value.trim();
  var membersVal = document.getElementById('f-members').value.trim();

  if (!dateVal || !dayVal || !timeVal || !roomVal || !groupVal || !phoneVal) {
    showAlert('err','⚠️ សូមបំពេញព័ត៌មានចាំបាច់: ថ្ងៃខែ, ម៉ោង, បន្ទប់, ឈ្មោះក្រុម, លេខទូរសព្ទ');
    return;
  }

  var d = new Date(dateVal);
  var formatted = ('0'+d.getDate()).slice(-2)+'/'+('0'+(d.getMonth()+1)).slice(-2)+'/'+d.getFullYear();

  var btn = document.getElementById('btn-submit');
  btn.disabled = true;
  btn.innerHTML = '<span class="loader"></span>កំពុងស្នើរសុំ...';

  var payload = {
    date: formatted, day: dayVal, time: timeVal,
    topic: topicVal||'—', group: groupVal,
    members: membersVal||'—', room: roomVal,
    phone: phoneVal, email: emailVal||'—'
  };

  fetch(SCRIPT_URL, {
    method: 'POST', mode: 'no-cors',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify(payload)
  })
  .then(function() {
    showAlert('ok','✅ <strong>ស្នើរសុំបានជោគជ័យ!</strong> បន្ទប់ '+roomVal+' ម៉ោង'+timeVal+' ('+formatted+')');
    clearForm();
    setTimeout(loadBookings, 1500);
  })
  .catch(function(err) {
    showAlert('err','❌ Error: '+err.message);
  })
  .finally(function() {
    btn.disabled = false;
    btn.innerHTML = '✅ ស្នើរសុំបន្ទប់';
  });
}

function clearForm() {
  ['f-date','f-day','f-time','f-room','f-topic','f-group','f-phone','f-email','f-members'].forEach(function(id){
    document.getElementById(id).value = '';
  });
}

var allBookings = [];

function loadBookings() {
  var btn = document.querySelector('.btn-refresh');
  btn.textContent = '⏳'; btn.disabled = true;

  fetch(SCRIPT_URL + '?action=get')
  .then(function(r){ return r.json(); })
  .then(function(data){
    allBookings = Array.isArray(data) ? data : [];
    applyFilter();
  })
  .catch(function(){
    document.getElementById('tbl-body').innerHTML =
      '<tr><td colspan="4"><div class="empty-state"><div class="icon">🔒</div>' +
      '<div>CORS blocked — ពិនិត្យ Apps Script "Who has access" = <strong>Anyone</strong></div></div></td></tr>';
  })
  .finally(function(){
    btn.textContent = '🔄 Refresh'; btn.disabled = false;
  });
}

function applyFilter() {
  var fr = document.getElementById('filter-room').value;
  var fs = document.getElementById('filter-status').value;
  var filtered = allBookings.filter(function(b){
    return (!fr||b.room===fr)&&(!fs||b.status===fs);
  });
  renderTable(filtered);
  document.getElementById('count-badge').textContent = filtered.length+' ការស្នើរសុំ';
}

function renderTable(rows) {
  var tb = document.getElementById('tbl-body');
  if (!rows||rows.length===0) {
    tb.innerHTML='<tr><td colspan="4"><div class="empty-state"><div class="icon">📭</div><div>មិនទាន់មានការស្នើរសុំ</div></div></td></tr>';
    return;
  }
  tb.innerHTML = rows.map(function(b){
    var badge = b.status==='កក់ហើយ'
      ? '<span class="badge badge-booked">✅ កក់ហើយ</span>'
      : '<span class="badge badge-free">🟡 ទំនេរ</span>';
    var dateStr = b.date ? (b.date.toString().includes('T') ? new Date(b.date).toLocaleDateString('km-KH') : b.date) : '—';
    return '<tr>'
      +'<td>'+dateStr+'</td>'
      +'<td>'+(b.time||'—')+'</td>'
      +'<td><span class="room-chip">'+(b.room||'—')+'</span></td>'
      +'<td>'+badge+'</td>'
      +'</tr>';
  }).join('');
}