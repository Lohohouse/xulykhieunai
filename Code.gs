/*************************************************************
 *  LOHO HOUSE - Backend Đơn xử lý khiếu nại (đa tab theo tháng)
 *  Mỗi tháng 1 sheet: T1.2026, T2.2026, ...
 *  Có: tự sinh mã, sửa mã (đổi tên / chuyển tab tháng), khóa ghi song song.
 *************************************************************/
var FIELDS = [
  ['maDon','Mã đơn'],['trangThai','Trạng thái'],['nguoiLapDon','Người lập đơn'],['donViLap','Đơn vị lập'],
  ['ngayLapDon','Ngày lập'],['gioLap','Giờ lập'],['tenKH','Tên khách hàng'],['sdt','Số điện thoại'],
  ['kenh','Kênh bán hàng'],['kenhKhac','Kênh khác'],['diaChi','Địa chỉ KH'],['loaiKN','Loại khiếu nại'],
  ['loaiKNKhac','Loại KN khác'],['maDonHang','Mã đơn hàng'],['maVanDon','Mã vận đơn'],['ngayDatHang','Ngày đặt hàng'],
  ['ngayDongHang','Ngày đóng hàng'],['sanPham','Sản phẩm KN'],['noiDung','Nội dung KN'],['yeuCau','Yêu cầu KH'],
  ['xuLyDeNghi','Cách xử lý đề nghị'],['nguoiXuLy','Người xử lý'],['chucVu','Chức vụ'],['boPhan','Bộ phận'],
  ['ngayNhanDon','Ngày nhận đơn'],['gioNhan','Giờ nhận'],['xacNhanKN','Xác nhận loại KN'],['xacNhanKhac','Xác nhận khác'],
  ['lyDo','Nêu rõ lý do'],['phuongAn','Phương án xử lý'],['kyChuQuan','Ký - Chủ quản'],['kyKeToan','Ký - Kế toán'],
  ['kyXuatKho','Ký - Xuất kho'],['kyCSKH','Ký - CSKH']
];
function monthTab_(maDon, ngayLapDon){
  var m=String(maDon||'').match(/KN(\d{2})(\d{2})/);
  if(m){ return 'T'+parseInt(m[1],10)+'.20'+m[2]; }
  var d=String(ngayLapDon||'').match(/^(\d{4})-(\d{2})/);
  if(d){ return 'T'+parseInt(d[2],10)+'.'+d[1]; }
  var now=new Date(); return 'T'+(now.getMonth()+1)+'.'+now.getFullYear();
}
function isDataTab_(name){ return /^T\d{1,2}\.\d{4}$/.test(name); }
function getSheetByName_(tab){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName(tab);
  if(!sh) sh=ss.insertSheet(tab);
  if(sh.getLastRow()===0){
    var header=['Thời gian ghi'].concat(FIELDS.map(function(f){return f[1];}));
    sh.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold').setBackground('#7a5c3e').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}
function findRow_(sh, maDon){
  if(!maDon) return -1; var last=sh.getLastRow(); if(last<2) return -1;
  var codes=sh.getRange(2,2,last-1,1).getValues();
  for(var i=0;i<codes.length;i++){ if(String(codes[i][0]).trim()===String(maDon).trim()) return i+2; }
  return -1;
}
/* Tìm 1 mã đơn ở bất kỳ tab tháng nào (dùng cho sửa mã / chuyển tab). */
function findAnyRow_(code){
  if(!code) return null;
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sheets=ss.getSheets();
  for(var s=0;s<sheets.length;s++){
    var sh=sheets[s]; if(!isDataTab_(sh.getName())) continue;
    var r=findRow_(sh,code); if(r>0) return {sheet:sh,row:r,tab:sh.getName()};
  }
  return null;
}
function nextCode_(){
  var now=new Date(), mm=('0'+(now.getMonth()+1)).slice(-2), yy=(''+now.getFullYear()).slice(-2);
  var prefix='KN'+mm+yy+'-', sh=getSheetByName_(monthTab_('KN'+mm+yy+'-001','')), max=0, last=sh.getLastRow();
  if(last>=2){ sh.getRange(2,2,last-1,1).getValues().forEach(function(r){
    var c=String(r[0]).trim(); if(c.indexOf(prefix)===0){ var n=parseInt(c.substring(prefix.length),10); if(!isNaN(n)&&n>max)max=n; } }); }
  return prefix+('00'+(max+1)).slice(-3);
}
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function fmtVal_(key,v){
  if(v instanceof Date){ var tz=SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    if(key==='gioLap'||key==='gioNhan') return Utilities.formatDate(v,tz,'HH:mm');
    return Utilities.formatDate(v,tz,'yyyy-MM-dd'); }
  return v;
}
function doGet(e){
  var action=(e&&e.parameter&&e.parameter.action)||'ping';
  if(action==='next') return json_({maDon:nextCode_()});
  if(action==='list'){
    var ss=SpreadsheetApp.getActiveSpreadsheet(), out=[];
    ss.getSheets().forEach(function(sh){
      if(!isDataTab_(sh.getName())) return;
      var last=sh.getLastRow(); if(last<2) return;
      var vals=sh.getRange(2,2,last-1,FIELDS.length).getValues();
      vals.forEach(function(row){ var obj={}; FIELDS.forEach(function(f,idx){ obj[f[0]]=fmtVal_(f[0],row[idx]); }); if(obj.maDon) out.push(obj); });
    });
    return json_(out);
  }
  return json_({ok:true,msg:'multi-tab'});
}
/* ============ TÀI KHOẢN & PHÂN QUYỀN ============ */
var USERS_SHEET='_TaiKhoan';
var DEFAULT_USERS=[['quanly','quanly@123','manager',1],['nhanvien','nhanvien@123','staff',1]];
function usersSheet_(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName(USERS_SHEET);
  if(!sh){ sh=ss.insertSheet(USERS_SHEET);
    sh.getRange(1,1,1,4).setValues([['username','password','role','mustChange']]);
    sh.getRange(2,1,DEFAULT_USERS.length,4).setValues(DEFAULT_USERS);
    try{ sh.hideSheet(); }catch(e){}
  }
  if(sh.getLastRow()<2){ sh.getRange(2,1,DEFAULT_USERS.length,4).setValues(DEFAULT_USERS); }
  return sh;
}
function findUser_(u){ var sh=usersSheet_(), last=sh.getLastRow(); if(last<2) return null;
  var v=sh.getRange(2,1,last-1,4).getValues();
  for(var i=0;i<v.length;i++){ if(String(v[i][0]).trim().toLowerCase()===String(u||'').trim().toLowerCase())
    return {row:i+2, username:String(v[i][0]), password:String(v[i][1]), role:String(v[i][2]), mustChange:(v[i][3]==1||v[i][3]==='1'||v[i][3]===true)}; }
  return null;
}

function doPost(e){
  var lock=LockService.getScriptLock(); try{ lock.waitLock(30000); }catch(le){}
  try{
    var payload=JSON.parse(e.postData.contents);
    var _act=payload.action;
    if(_act==='login'){
      var lu=findUser_(payload.username);
      if(lu && lu.password===String(payload.password)) return json_({ok:true, role:lu.role, mustChange:lu.mustChange, username:lu.username});
      return json_({ok:false, error:'Sai tài khoản hoặc mật khẩu'});
    }
    if(_act==='setpw'){
      var su=findUser_(payload.username);
      if(!su || su.password!==String(payload.oldPassword)) return json_({ok:false, error:'Mật khẩu hiện tại không đúng'});
      var np=String(payload.newPassword||''); if(np.length<4) return json_({ok:false, error:'Mật khẩu mới tối thiểu 4 ký tự'});
      var ush=usersSheet_(); ush.getRange(su.row,2).setValue(np); ush.getRange(su.row,4).setValue(0);
      return json_({ok:true});
    }
    if(_act==='delete'){
      var dl=findAnyRow_(String(payload.maDon||'').trim());
      if(dl){ dl.sheet.deleteRow(dl.row); return json_({ok:true, action:'delete'}); }
      return json_({ok:false, error:'Không tìm thấy đơn'});
    }
    if(payload.action!=='save') return json_({ok:false,error:'action khong hop le'});
    var data=payload.data||{};
    var maDon=String(data.maDon||'').trim();
    if(!maDon || maDon.indexOf('(')>=0) maDon=nextCode_();
    data.maDon=maDon;
    var oldCode=String(data.maDonGoc||'').trim();        // mã gốc trước khi sửa (nếu có)
    var rowVals=[new Date()].concat(FIELDS.map(function(f){return data[f[0]]||'';}));
    var newTab=monthTab_(maDon, data.ngayLapDon);
    var newSheet=getSheetByName_(newTab);
    // ưu tiên tìm theo mã gốc (để sửa mã); nếu không có mã gốc thì tìm theo mã mới
    var loc=findAnyRow_(oldCode||maDon);
    var action;
    if(loc){
      if(loc.tab===newTab){
        newSheet.getRange(loc.row,1,1,rowVals.length).setValues([rowVals]); action='update';
      } else {
        loc.sheet.deleteRow(loc.row);        // chuyển tháng: xóa dòng cũ, thêm vào tab mới
        newSheet.appendRow(rowVals); action='move';
      }
    } else {
      var ex=findRow_(newSheet,maDon);
      if(ex>0){ newSheet.getRange(ex,1,1,rowVals.length).setValues([rowVals]); action='update'; }
      else { newSheet.appendRow(rowVals); action='insert'; }
    }
    return json_({ok:true,maDon:maDon,action:action});
  }catch(err){ return json_({ok:false,error:String(err)}); }
  finally{ try{ lock.releaseLock(); }catch(le2){} }
}
/* Sắp xếp lại tất cả tab tháng theo Mã đơn (cột B) tăng dần.
   Chạy thủ công trong Apps Script khi cần dồn thứ tự. */
function sortAll(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function(sh){
    if(!/^T\d{1,2}\.\d{4}$/.test(sh.getName())) return;
    var last=sh.getLastRow(); if(last<3) return;
    sh.getRange(2,1,last-1,sh.getLastColumn()).sort({column:2,ascending:true});
  });
}
