function exec() {
  const reservationInfoFromGmail = getReservationInfoFromGmail()
  exportToSheet(reservationInfoFromGmail, 'reservation')

  Utilities.sleep(1000)

  const cancelInfoFromGmail = getCancelInfoFromGmail()
  exportToSheet(cancelInfoFromGmail, 'cancel')
}
