function exportToSheet(infoFromGmail, actionKind) {
  const RESERVATION = 'reservation'
  const CANCEL = 'cancel'
  const GRAY_COLOUR = "#afafb0"

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NAME)
  // 3行目から、A列の値(予約No)を全件取得して配列に格納する
  // const existReservationIds = sheet.getRange(3, 1, sheet.getLastRow() - 1).getValues().flat()

  if (actionKind === RESERVATION) {
    for(gmailInfo of infoFromGmail) {
      const existReservationIds = sheet.getRange(3, 1, sheet.getLastRow() - 1).getValues().flat()
      const reservationId = gmailInfo[0]
      if (existReservationIds.includes(reservationId)) {continue} //既にスプシに記入されてる場合はスキップ

      sheet.appendRow(gmailInfo)
      protectLastRow(sheet)
    }
  }

  if (actionKind === CANCEL) {
    for(gmailInfo of infoFromGmail) {
      const existReservationIds = sheet.getRange(3, 1, sheet.getLastRow() - 1).getValues().flat()
      const cancelledReservationId = gmailInfo.reservationId
      const row = getRowByKeyword(cancelledReservationId, existReservationIds)
      // キャンセルされた予約番号が存在しない場合は、対象をログに残してスキップ
      if (row === 0) {
        Logger.log(`予約番号「${cancelledReservationId}」のキャンセルメールがあるが、シートに該当の予約番号が見つかりません。`)
        continue
      }

      sheet.getRange(row, 18).setValue(`GAS：予約番号「${cancelledReservationId}」のご予約のキャンセルメールがあります。`)
      sheet.getRange(`A${row}:R${row}`).setBackground(GRAY_COLOUR)
    }
  }
}

function getRowByKeyword(keyword, existReservationIds) {
  const row = existReservationIds.indexOf(keyword) + 1
  // indexOfメソッドはkeywordが見つからない場合に-1を返す。つまりrow === 0は、keywordが見つからない場合。
  if (row === 0) {return 0}

  // 上2行はヘッダーのため+2している
  return row + 2
}

function protectLastRow(sheet) {
  const lastRow = sheet.getLastRow()
  const rangesToProtect =  [
    `A${lastRow}:G${lastRow}`,  // 最終行(今insertした行の)A-G列を保護
    `I${lastRow}:J${lastRow}`,  // つまり、H,K,O,Q,R列「以外」を保護
    `L${lastRow}:N${lastRow}`,　
    `P${lastRow}`,
  ]

  for(range of rangesToProtect) {
    const  protectedRange = sheet.getRange(range).protect()
    //保護したシートで編集可能なユーザーを取得
    const userList = protectedRange.getEditors()
    //オーナーのみ編集可能にするため、編集ユーザーをすべて削除
    //オーナーの編集権限は削除できないため、オーナーのみ編集可能に
    protectedRange.removeEditors(userList)
    protectedRange.setDescription(`GAS：${range}を保護`);
  }
}
