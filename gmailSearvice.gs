// 予約メール取得 ======================================================================================================

function getReservationInfoFromGmail() {
  const jaranReservationInfo = getJalanReservationInfo()
  const skyReservationTicketInfo = getSkyTicketReservationInfo()
  const mergedReservationArr = jaranReservationInfo.concat(skyReservationTicketInfo)

  return mergedReservationArr
}

function getJalanReservationInfo() {
  const query = 'subject:じゃらんnetレンタカー 予約通知';
  const threads = GmailApp.search(query);
  const messagesForThreads = GmailApp.getMessagesForThreads(threads)
  let gmailInfo = []
  for(const messages of messagesForThreads){
    for(const message of messages) {
      if (message.getSubject().match(/Re:(.*)/)) {continue} // 件名に「Re:」の文字列がある場合はスキップ

      const body = message.getBody()
      // .replace(/.+?：/, '')は、「先頭の文字列から：まで」を削除する正規表現
      // ex) '予約番号： R0VONPA2' → 'R0VONPA2'
      //
      // .trim()は、文字列前後の空白を削除している(文字列内の空白は削除しない)
      const reservationId = body.match(/予約番号：(.*)/)[0].replace(/.+?：/, '').trim()
      const email = body.match(/予約者メールアドレス：(.*)/)[0].replace(/.+?：/, '').trim()
      const driverName = body.match(/運転者氏名：(.*)/)[0].replace(/.+?：/, '').trim()
      const driverNameKana = body.match(/運転者氏名カナ：(.*)/)[0].replace(/.+?：/, '').trim()
      const startDate = body.match(/貸出日時：(.*)/)[0].replace(/.+?：/, '').trim().split(' ')[0]
      const startTime = body.match(/貸出日時：(.*)/)[0].replace(/.+?：/, '').trim().split(' ')[1]
      const returnDate = body.match(/返却日時：(.*)/)[0].replace(/.+?：/, '').trim().split(' ')[0]
      const returnTime = body.match(/返却日時：(.*)/)[0].replace(/.+?：/, '').trim().split(' ')[1]
      const selectedPlan = body.match(/料金プラン：(.*)/)[0].replace(/.+?：/, '').trim()
      const options = body.match(/オプション：(.*)/)[0].replace(/.+?：/, '').trim()
      const passengers = body.match(/乗車人数：(.*)/)[0].replace(/.+?：/, '').trim()
      const totalPrice = body.match(/利用者への請求額：(.*)/)[0].replace(/.+?：/, '').trim()
      // じゃらんの手数料は入金額の12%
      // totalPriceが'73,980円'のような文字列のため、カンマと円を削除してから計算してる
      const charge = Number(totalPrice.replace(/[^0-9]/g, '')) * 0.12
      const totalAfterChargeWithFormat = (Number(totalPrice.replace(/[^0-9]/g, '')) - charge).toLocaleString() + '円'

      gmailInfo.push([
        reservationId,
        driverName + ' ' + driverNameKana,
        email,
        startDate,
        startTime,
        returnDate,
        returnTime,
        '', //H列：免許証情報を手動選択するため空白
        passengers,
        selectedPlan,
        '', //K列：車種を手動選択するため空白
        options,
        totalPrice,
        '', //N列：じゃらんはメール予約してから入金手続きをするため、この時点では空白(スカイチケットはメール予約と同時に入金だから空白ではない)
        '', //O列：手動入力の箇所なので空白
        totalAfterChargeWithFormat
      ])
    }
  }

  return gmailInfo
}

function getSkyTicketReservationInfo() {
  const query = 'subject:【skyticket】新規予約';
  const threads = GmailApp.search(query);
  const messagesForThreads = GmailApp.getMessagesForThreads(threads)
  let gmailInfo = []

  for(const messages of messagesForThreads){
    for(const message of messages) {
      if (message.getSubject().match(/Re:(.*)/)) {continue} // 件名に「Re:」の文字列がある場合はスキップ

      const body = message.getBody()
      const reservationId = body.match(/予約番号：(.*)/)[0].replace(/.+?：/, '').trim()
      const driverName = body.match(/ご利用者名：(.*)/)[0].replace(/.+?：/, '').trim()
      const email = body.match(/メールアドレス：(.*)/)[0].replace(/.+?：/, '').trim()
      //'2023年02月02日(木)'のように(木)があるとスプシで日時のソートができなくなるので、(木)の部分を削除する
      const startDate = body.match(/受取日時：(.*)/)[0].replace(/.+?：/, '').trim().split(' ')[0].replace(/(\(.*\))/, '')
      const startTime = body.match(/受取日時：(.*)/)[0].replace(/.+?：/, '').trim().split(' ')[1]
      //'2023年02月02日(木)'のように(木)があるとスプシで日時のソートができなくなるので、(木)の部分を削除する
      const returnDate = body.match(/返却日時：(.*)/)[0].replace(/.+?：/, '').trim().split(' ')[0].replace(/(\(.*\))/, '')
      const returnTime = body.match(/返却日時：(.*)/)[0].replace(/.+?：/, '').trim().split(' ')[1]
      const passengers = body.match(/ご利用人数：(.*)/)[0].replace(/.+?：/, '').trim()
      const selectedPlan = body.match(/プラン名：(.*)/)[0].replace(/.+?：/, '').trim()
      let options
      if (body.match(/オプション：(.*)/)) {
        options = body.match(/オプション：(.*)/)[0].replace(/.+?：/, '').trim()
      } else {
        options = ''
      }
      const totalPrice = body.match(/合計料金：(.*)/)[0].replace(/.+?：/, '').trim()
      const isPaid = body.match(/入金状況：(.*)/)[0].replace(/.+?：/, '').trim() === '入金済み' ? true : false
      let paymentStatus
      let paidAmount
      if (isPaid) {
        paymentStatus = '済'
        paidAmount = body.match(/入金金額：(.*)/)[0].replace(/.+?：/, '').trim()
      } else {
        paymentStatus = '未'
      }
      // スカイチケットの手数料は入金額の15%
      // totalPriceが'73,980円'のような文字列のため、カンマと円を削除してから計算してる
      const charge = Number(totalPrice.replace(/[^0-9]/g, '')) * 0.15
      const totalAfterChargeWithFormat = (Number(totalPrice.replace(/[^0-9]/g, '')) - charge).toLocaleString() + '円'

      gmailInfo.push([
        reservationId,
        driverName,
        email,
        startDate,
        startTime,
        returnDate,
        returnTime,
        '', //H列：免許証情報を手動選択するため空白
        passengers,
        selectedPlan,
        '', //K列：車種を手動選択するため空白
        options,
        totalPrice,
        paidAmount || '',
        '', //O列：手動入力の箇所なので空白
        totalAfterChargeWithFormat,
        paymentStatus,
      ])
    }
  }

  return gmailInfo
}

// キャンセルメール取得 ======================================================================================================

function getCancelInfoFromGmail() {
  const jaranCancelInfo = getJaranCancelInfo()
  const skyTicketCancelInfo = getSkyTicketCancelInfo()
  const mergedCancelInfoArr = jaranCancelInfo.concat(skyTicketCancelInfo)

  return mergedCancelInfoArr
}

function getJaranCancelInfo() {
  const query = 'subject:じゃらんnetレンタカー キャンセル通知';
  const threads = GmailApp.search(query);
  const messagesForThreads = GmailApp.getMessagesForThreads(threads)

  let cancelInfo = []
  for(const messages of messagesForThreads){
    for(const message of messages) {
      if (message.getSubject().match(/Re:(.*)/)) {continue} // 件名に「Re:」の文字列がある場合はスキップ

      const contactedDate = message.getDate()
      const body = message.getBody()
      const reservationId = body.match(/予約番号：(.*)/)[0].replace(/.+?：/, '').trim()
      // const cancellationDateTime = body.match(/キャンセル受付時間：(.*)/)[0].replace(/.+?：/, '').replace('）', '').trim()

      cancelInfo.push({
        reservationId: reservationId,
        contactedDate: contactedDate
      })
    }
  }

  return cancelInfo
}

function getSkyTicketCancelInfo() {
  const query = 'subject:【skyticket】キャンセル';
  const threads = GmailApp.search(query);
  const messagesForThreads = GmailApp.getMessagesForThreads(threads)

  let cancelInfo = []
  for(const messages of messagesForThreads){
    for(const message of messages) {
      if (message.getSubject().match(/Re:(.*)/)) {continue} // 件名に「Re:」の文字列がある場合はスキップ

      const contactedDate = message.getDate()
      const body = message.getBody()
      const reservationId = body.match(/予約番号：(.*)/)[0].replace(/.+?：/, '').trim()
      // const cancellationDateTime = body.match(/キャンセル受付時間：(.*)/)[0].replace(/.+?：/, '').replace('）', '').trim()

      cancelInfo.push({
        reservationId: reservationId,
        contactedDate: contactedDate
      })
    }
  }

  return cancelInfo
}
