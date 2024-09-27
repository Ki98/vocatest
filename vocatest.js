let spread;

window.onload = function () {
    spread = new GC.Spread.Sheets.Workbook(document.getElementById("ss"));

    // 초기 데이터 로드 후 이벤트 리스너 추가
    document.querySelectorAll('input[name="chk_info"]').forEach((radio) => {
        radio.addEventListener('change', hideColumnsByLanguage);
    });
};

function ImportFile() {
    let file = document.getElementById("fileDemo").files[0];
    let fileType = file.name.split('.');
    if(fileType[fileType.length - 1] == 'xlsx') {
        spread.import(file, function () {
            // success callback
        }, function (e) {
            console.log(e); // error callback
        }, {
            fileType: GC.Spread.Sheets.FileType.excel
        });
    }
}

function Export_Excel() {
    let fileName = document.getElementById("exportFileName").value;
    if (fileName.substr(-5, 5) !== '.xlsx') {
        fileName += '.xlsx';
    }
    spread.export(function (blob) {
        saveAs(blob, fileName);
    }, function (e) {
        console.log(e);
    }, { 
        fileType: GC.Spread.Sheets.FileType.excel
    });
}

function hideColumnsByLanguage() {
    const sheet = spread.getActiveSheet();
    const columnCount = sheet.getColumnCount();
    const selectedOption = document.querySelector('input[name="chk_info"]:checked').value;

    // 모든 행을 먼저 보이게 함
    for (let i = 0; i < columnCount; i++) {
        sheet.setColumnVisible(i, true); // 모든 행 보이게
    }

    if (selectedOption === "blind_english") {
        // 홀수 행 숨기기 (1, 3, 5,... 인덱스 기준)
        for (let i = 0; i < columnCount; i++) {
            if (i % 2 === 0) { // 짝수 행
                sheet.setColumnVisible(i, false);
            }
        }
    } else if (selectedOption === "blind_korean") {
        // 짝수 행 숨기기 (0, 2, 4,... 인덱스 기준)
        for (let i = 0; i < columnCount; i++) {
            if (i % 2 === 1) { // 홀수 행
                sheet.setColumnVisible(i, false);
            }
        }
    }
}

function printExcel() {
    console.log(spread); // spread 객체 확인
    if (typeof spread.print === 'function') {
        spread.print();
    } else {
        console.error("print 메서드가 없습니다.");
    }
}