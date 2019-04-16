function loadScript(url, callback) {
  const head = document.head;
  const script = document.createElement('script');
  script.type = 'text/javascript';
  script.src = url;
  script.onreadystatechange = callback;
  script.onload = callback;
  head.appendChild(script);
}
loadScript("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.14.2/xlsx.core.min.js", () => {
  $("html").prepend(`
    <label for="avatar">Choose a file:</label>
    <input type="file" id="file">
    <select id="sheet" disabled></select>
    <input type="text" id="stt" />
  `);
})
let configAddStaff = [
  {
    "name": "Công ty",
    "input": "_ctl0:cboCompanyID",
    dropdown: true
  },
  {
    "name": "Mã nhân viên",
    "input": "_ctl0:txtMaNhanVien"
  },
  {
    "name": "Khối",
    "input": "_ctl0:cboLevel1ID"
  },
  {
    "name": "Trung tâm",
    "input": "_ctl0:cboLevel2ID"
  },
  {
    "name": "Phòng",
    "input": "_ctl0:cboLevel3ID"
  },
  {
    "name": "Bộ phận",
    "input": "_ctl0:cboLSLevel4IDAll"
  },
  {
    "name": "Nhóm",
    "input": "_ctl0:cboLSLevel5IDAll"
  },
  {
    "name": "Họ",
    "input": "_ctl0:txtHoTenLot"
  },
  {
    "name": "Tên đệm",
    "input": "_ctl0:txtMiddleName"
  },
  {
    "name": "Tên gọi",
    "input": "_ctl0:txtTen"
  },
  {
    "name": "Nơi sinh (Text)",
    "input": "_ctl0:txtNoiSinh"
  },
  {
    "name": "Nơi sinh",
    "input": "_ctl0:cboNoiSinh_LSProvinceID"
  },
  {
    "name": "Giới tính",
    "input": "_ctl0:cboGioiTinh"
  },
  {
    "name": "Tình trạng hôn nhân",
    "input": "_ctl0:cboTinhTrangHonNhan"
  },
  {
    "name": "Quốc tịch",
    "input": "_ctl0:cboLSNationalityID",
    dropdown: true
  },
  {
    "name": "Dân tộc",
    "input": "_ctl0:cboLSEthnicID",
    dropdown: true
  },
  {
    "name": "Nơi cấp CMND",
    "input": "_ctl0:cboNoiCapCMND"
  },
  {
    "name": "Nhập Passport",
    "input": "_ctl0:chkPassport"
  },
  {
    "name": "Số passport",
    "input": "_ctl0:txtPassportNo"
  },
  {
    "name": "Ngày cấp passport",
    "input": "_ctl0:txtNgayCapPassport"
  },
  {
    "name": "Ngày cấp CMND",
    "input": "_ctl0:txtNgayCapCMND"
  },
  {
    "name": "CMND Số",
    "input": "_ctl0:txtSoCMND"
  },
  {
    "name": "Ngày sinh",
    "input": "_ctl0:txtNgaySinh"
  },
  {
    "name": "Ngày hiệu lực",
    "input": "_ctl0:txtNgayHieuLucPassport"
  },
  {
    "name": "Ngày hết hạn",
    "input": "_ctl0:txtNgayHetHanPassport"
  },
  {
    "name": "Loại Passport",
    "input": "_ctl0:cboLoaiPassport"
  },
  {
    "name": "Nơi cấp Passport",
    "input": "_ctl0:txtNoiCapPassport"
  },
  {
    "name": "Nơi cấp",
    "input": "_ctl0:cboNoiCapMST"
  },
  {
    "name": "Ngày vào chính thức",
    "input": "_ctl0:txtNgayVaoChinhThuc"
  },
  {
    "name": "Loại nhân viên",
    "input": "_ctl0:cboLoaiNhanVien"
  },
  {
    "name": "Chức danh",
    "input": "_ctl0:cboLSJobTitleID_Related"
  },
  {
    "name": "Trình độ chuyên môn",
    "input": "_ctl0:cboTrinhDoChuyenMon"
  },
  {
    "name": "Số di động",
    "input": "_ctl0:txtSoDiDong"
  },
  {
    "name": "Chức vụ",
    "input": "_ctl0:cboLSChucVuID"
  },
  {
    "name": "Nơi Làm việc",
    "input": "_ctl0:cboLocationID"
  },
  {
    "name": "Ghi chú",
    "input": "_ctl0:txtGhiChuPassport"
  },
  {
    "name": "Ngày vào Công ty",
    "input": "_ctl0:txtNgayVaoCongTy"
  },
  {
    "name": "Hình thức thay đổi",
    "input": "_ctl0:cboLSStatusChangeID"
  },
  {
    "name": "Trình độ học vấn",
    "input": "_ctl0:cboTrinhDoHocVan"
  },
  {
    "name": "Nhóm chấm công",
    "input": "_ctl0:cboLoaiHinhNhanVien"
  },
  {
    "name": "Local/Expat",
    "input": "_ctl0:cboNhomNhanVien"
  },
  {
    "name": "Cấp trên gián tiếp",
    "input": "_ctl0:txtEmpIndirectReport"
  },
  {
    "name": "Số di động",
    "input": "_ctl0:txtSoDiDong"
  },
  {
    "name": "Email cá nhân",
    "input": "_ctl0:txtEmail"
  },
  {
    "name": "Email công ty",
    "input": "_ctl0:txtCompanyEmail"
  },
  {
    "name": "Mã số thuế",
    "input": "_ctl0:txtMaSoThue"
  },
  {
    "name": "Ngày cấp",
    "input": "_ctl0:txtNgayCapMST"
  },
  {
    "name": "Nơi cấp",
    "input": "_ctl0:cboNoiCapMST"
  },
  {
    "name": "Địa chỉ thường trú",
    "input": "_ctl0:txtDiaChiThuongTru"
  },
  {
    "name": "Địa chỉ tạm trú",
    "input": "_ctl0:txtDiaChiTamTru"
  },
  {
    "name": "Địa chỉ CMND",
    "input": "_ctl0:txtDiaChiTamTruEN"
  },
  {
    "name": "Nơi sinh",
    "input": "_ctl0:txt_NoiSinh"
  },
  {
    "name": "Ghi chú",
    "input": "_ctl0:txtNoteAdd"
  }
];
let WK_GLOBAL = {};
$(document).ready(function (e) {

  function to_json(workbook) {
    var result = {};
    workbook.SheetNames.forEach(function (sheetName) {
      var roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
      if (roa.length > 0) {
        result[sheetName] = roa;
      }
    });
    return result;
  }

  function handleFile(e) {
    var files = e.target.files;
    var i, f;
    for (i = 0, f = files[i]; i != files.length; ++i) {
      var reader = new FileReader();
      var name = f.name;
      reader.onload = function (e) {
        try {
          var data = $.trim(e.target.result);
          var workbook = XLSX.read(data, {
            type: 'binary'
          });
          const wkJson = to_json(workbook);
          WK_GLOBAL = wkJson;
          fillSheet(wkJson);
        } catch (e) {
          alert(e);
        }
      };
      reader.readAsBinaryString(f);
    }
  }

  function setData(wk, sheet, stt) {
    const current = wk[sheet][stt];
    $('#_ctl0_chkPassport').click();
    handle(current);
  }

  function fillSheet(wk) {
    const arrSheet = Object.keys(wk);
    $('#sheet').html(arrSheet.map(e => `<option value="${e}">${e}</option>`))
  }

  $("html")
    .on('change', '#stt', function () {
      setData(WK_GLOBAL, $("#sheet").val(), $(this).val())
    })
    .on('change', '#file', function (e) {
      e.preventDefault();
      handleFile(e);
    })
});

function eventFire(el, etype) {
  if (el.fireEvent) {
    el.fireEvent('on' + etype);
  } else {
    const evObj = document.createEvent('Events');
    evObj.initEvent(etype, true, false);
    el.dispatchEvent(evObj);
  }
}

function handleSetData(idInput, val) {
  return new Promise((resolve, reject) => {
    const idTarget = `${idInput}_DropDown`;
    const qTarget = `#Form1 .rcbSlide #${idTarget} li`;
    if ($(qTarget).length > 0) {
      return resolve(setDataDropdown(qTarget, val));
    }
    return $("form").bind("DOMSubtreeModified", function () {
      setTimeout(() => {
        console.log($(qTarget).length);
        console.log($(qTarget).html())
        $("form").off('DOMSubtreeModified');
        if ($(qTarget).length > 0) {
          return resolve(setDataDropdown(qTarget, val));
        }
        return resolve(false);
      }, 1000)
    });
  });
}

function setDataDropdown(qEls, val) {
  let setSuccessfully = false;
  $(qEls).each(function () {
    if ($(this).text() === val) {
      setSuccessfully = true;
      $(this).click();
    }
  })
  return setSuccessfully;
}

async function handle(current) {
  const arr = Object.keys(current).map(e => {
    const c = configAddStaff.find(a => a.name.toLocaleLowerCase() === e.toLocaleLowerCase());
    if (c) {
      return {
        input: c.input,
        value: current[e],
        dropdown: c.dropdown
      }
    }
    return null;
  }).filter(e => e);
  const runFbyAwait = async (index) => {
    if (index > arr.length - 1) {
      return;
    }
    const currentData = arr[index];
    if (currentData.dropdown) {
      await inputDropdown(currentData.input.replace(":", "_"), currentData.value);
    } else {
      inputNormal(currentData.input, currentData.value)
    }

    return runFbyAwait(++index);
  }
  return await runFbyAwait(0);
}

async function inputDropdown(id, val) {
  eventFire($(`#${id}_Input`)[0], 'focus');
  console.log("Handle input dropdown " + id + " value: " + val);
  if (await handleSetData(id, val)) {
    console.log("Set data successfully")
  } else {
    console.log("Unset data")
  }
  eventFire($('body')[0], 'click');
  console.log("\n");
  return;
}

function inputNormal(id, val) {
  return $(`input[name='${id}']`).val(val);
}
