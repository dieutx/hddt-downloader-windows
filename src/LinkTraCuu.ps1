# Bảng tra cứu hóa đơn điện tử, chuyển từ sheet LinkTraCuu của
# TaiHoaDonDienTu.  Đây là dữ liệu tham chiếu công khai: MST của nhà cung
# cấp dịch vụ T-VAN và địa chỉ trang tra cứu của nhà cung cấp đó.
#
# Sheet này vừa là nguồn dữ liệu tham chiếu, vừa được ghi ra thành sheet
# 'LinkTraCuu' trong workbook xuất để người dùng tự thêm/sửa link.  Cách dùng
# bảng trong VBA gốc:
#   LinkTraCuu()    -> dicLink(cột B) = cột D   (MST nhà cung cấp -> link)
#   tenCotTraCuu()  -> dicTenCotTC(cột E)       (tên trường chứa mã tra cứu)
#   frmTaiHoaDon    -> arrTrangThai = I2:I8, arrKQKTHoaDon = O2:O11
#   modGhiExcel.bas -> cột 55 (Link tra cứu), 56 (Mã tra cứu) của
#                      TongHopHD_Mua/TongHopHD_Ban lấy từ dicLink/dicTenCotTC
# Khi một MST có nhiều dòng thì dòng cuối thắng, đúng như VBA gán đè vào
# dictionary.

# Cột A..O của sheet LinkTraCuu; phần tử đầu là dòng tiêu đề.
$script:TaiHoaDonDienTuLookupColumns = @('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O')
$script:TaiHoaDonDienTuLookupSheetData = @(
    [ordered]@{ A = 'Tên tổ chức'; B = 'MST'; C = 'MST Người bán'; D = 'Link tra cứu'; E = 'Tên trường mã tra cứu'; F = 'Ghi chú'; H = 'tthai'; I = 'Trạng thái hóa đơn MUA'; K = 'ttxly'; L = 'Kết quả kiểm tra MUA'; N = 'ttxly'; O = 'Kết quả kiểm tra BAN' },
    [ordered]@{ A = 'CÔNG TY ĐIỆN LỰC BẮC GIANG'; B = '0100100417-007'; D = 'https://bill.payoo.vn/tra-tien-thanh-toan-hoa-don-dien-evn?AspxAutoDetectCookieSupport=1'; E = 'TenDviQly'; H = 'All'; I = 'Tất cả'; K = 'All'; L = 'Tất cả'; N = 'All'; O = 'Tất cả' },
    [ordered]@{ A = 'Tập đoàn Công nghiệp - Viễn thông quân đội'; B = '0100109106'; C = '0109266456'; D = 'https://giaothongso.com.vn/tra-cuu-hoa-don-mtc/'; E = 'Mã số bí mật'; H = '1'; I = 'Hóa đơn mới'; K = '5'; L = 'Đã cấp mã hóa đơn'; N = '0'; O = 'Tổng cục thuế đã nhận' },
    [ordered]@{ A = 'Tập đoàn Công nghiệp - Viễn thông quân đội'; B = '0100109106'; D = 'https://vinvoice.viettel.vn/utilities/invoice-search'; E = 'Mã số bí mật'; H = '2'; I = 'Hóa đơn thay thế'; K = '6'; L = 'Tổng cục thuế đã nhận không mã'; N = '1'; O = 'Đang tiến hành kiểm tra điều kiện cấp mã' },
    [ordered]@{ A = 'Tập đoàn Công nghiệp - Viễn thông quân đội'; B = '0100109106'; C = '0107766421'; D = 'https://tracuuhoadon.vetc.com.vn/'; E = 'reservationCode'; H = '3'; I = 'Hóa đơn điều chỉnh'; K = '8'; L = 'Tổng cục thuế đã nhận hóa đơn có mã khởi tạo từ máy tính tiền'; N = '2'; O = 'CQT từ chối theo từng lần phát sinh' },
    [ordered]@{ A = 'Tập Đoàn Bưu chính viễn thông Việt Nam'; B = '0100684378'; C = '0108005532'; D = 'https://hoadon.petrolimex.com.vn/'; E = 'Fkey'; H = '4'; I = 'Hóa đơn bị thay thế'; N = '3'; O = 'Hóa đơn đủ điều kiện cấp mã' },
    [ordered]@{ A = 'Tập Đoàn Bưu chính viễn thông Việt Nam'; B = '0100684378'; C = '0101245486'; D = 'https://e-invoice-tt78.vingroup.net/'; E = 'Fkey'; H = '5'; I = 'Hóa đơn đã bị điều chỉnh'; N = '4'; O = 'Hóa đơn không đủ điều kiện cấp mã' },
    [ordered]@{ A = 'Tập Đoàn Bưu chính viễn thông Việt Nam'; B = '0100684378'; D = 'https://portaltool-miennam.vnpt-invoice.com.vn/'; E = 'Fkey'; H = '6'; I = 'Hóa đơn bị hủy'; N = '5'; O = 'Đã cấp mã hóa đơn' },
    [ordered]@{ A = 'Tổng công ty Viễn thông Mobifone'; B = '0100686209'; D = 'http://tracuuhoadon.mobifoneinvoice.vn/trang-chu'; N = '6'; O = 'Tổng cục thuế đã nhận không mã' },
    [ordered]@{ A = 'Công ty Cổ phần dịch vụ viễn thông và in Bưu điện'; B = '0100687474'; D = 'https://hoadondientu-ptp.vn/tra-cuu/'; H = 'tchat'; I = 'Tính chất'; N = '7'; O = 'Đã kiểm tra HĐĐT định kỳ không có mã' },
    [ordered]@{ A = 'Công ty Cổ phần phần mềm quản lý doanh nghiệp'; B = '0100727825'; D = 'https://einvoice.fast.com.vn/'; E = 'KeySearch'; H = '1'; I = 'Hàng hóa dịch vu'; N = '8'; O = 'Tổng cục thuế đã nhận hóa đơn có mã khởi tạo từ máy tính tiền' },
    [ordered]@{ A = 'Công ty Cổ phần Phần mềm Thăng Long'; B = '0101010702'; H = '2'; I = 'Khuyến mãi' },
    [ordered]@{ A = 'Công ty Cổ phần phát triển phần mềm ASIA'; B = '0101162173'; D = 'https://asiainvoice.vn/tra-cuu'; H = '3'; I = 'Chiết khấu thương mãi (theo dòng)'; O = 'Các tên field cho [tthue] và ThTiencoVAT' },
    [ordered]@{ A = 'Công ty Cổ phần MISA'; B = '0101243150'; D = 'https://www.meinvoice.vn/tra-cuu/'; E = 'TransactionID'; H = '4'; I = 'Ghi chú'; N = 'TThueVAT'; O = 'Tiền thuế,VATAmount,TongTien_Thue,Tiền thuế dòng (Tiền thuế GTGT)' },
    [ordered]@{ A = 'Công ty TNHH Phần mềm Nhân Hòa'; B = '0101289966'; D = 'https://tracuu.e-hoadon.cloud/'; N = 'THTiencoVAT'; O = 'Thành tiền có thuế, Amount,TongTien_CoThue,Thành tiền thanh toán của hàng hóa' },
    [ordered]@{ A = 'Công ty TNHH Phát triển công nghệ Thái Sơn'; B = '0101300842'; D = 'https://einvoice.vn/tra-cuu'; E = 'MaTC'; H = 'khmshdon' },
    [ordered]@{ A = 'Công ty Cổ phần Giải pháp hóa đơn điện tử Việt Nam'; B = '0101352495'; D = 'https://tracuu.v50.vninvoice.vn/'; H = '1'; I = 'Hóa đơn giá trị gia tăng' },
    [ordered]@{ A = 'Công ty Cổ phần BKAV'; B = '0101360697'; D = 'https://van.ehoadon.vn/Lookup?InvoiceGUID='; F = 'Lấy mã hóa đơn <DLHDon Id>'; H = '2'; I = 'Hóa đơn bán hàng' },
    [ordered]@{ A = 'Công ty Cổ phần Công nghệ và giải pháp Tâm Việt'; B = '0101622374'; H = '3'; I = 'Phiếu xuất kho kiêm vận chuyển' },
    [ordered]@{ A = 'Công ty Cổ phần GMO-Z.com RUNSYSTEM'; B = '0101659906'; D = 'https://tracuu.kaike.vn/-/'; H = '4'; I = 'Hóa đơn khác' },
    [ordered]@{ A = 'Công ty TNHH Tổng công ty Công nghệ và Giải pháp CMC'; B = '0101925883' },
    [ordered]@{ A = 'Công ty Cổ phần giải pháp thanh toán Việt Nam'; B = '0102182292'; D = 'https://einvoice.vnpay.vn/' },
    [ordered]@{ A = 'Công ty Cổ phần Công nghệ thông tin Đông Nam Á'; B = '0102454468'; D = 'https://tax24.com.vn/thuedientu/xac-minh-hoa-don'; E = 'Mã Kiểm tra' },
    [ordered]@{ A = 'Công ty Cổ phần Công nghệ tin học EFY Việt Nam'; B = '0102519041'; D = 'https://ihoadon.vn/kiem-tra/?lang=vn' },
    [ordered]@{ A = 'Công ty Cổ phần Hóa đơn điện tử TIG Thăng Long'; B = '0102720409' },
    [ordered]@{ A = 'Trung tâm Tin học và Công nghệ số'; B = '0102723181' },
    [ordered]@{ A = 'Công ty Cổ phần tích hợp công nghệ VNISC'; B = '0103018807'; D = 'https://abcsys.vn/Invoice/Search' },
    [ordered]@{ A = 'Công ty Cổ phần Tin học - Viễn thông Hàng không'; B = '0103019524'; D = 'https://einvoice.aits.vn/' },
    [ordered]@{ A = 'Công ty Cổ phần phát triển phần mềm và công nghệ Bitware'; B = '0103770970'; D = 'https://www.bitware.vn/tracuuhoadon/' },
    [ordered]@{ A = 'Công ty Cổ phần công nghệ thẻ Nacencomm'; B = '0103930279'; D = 'https://hoadon78_logigo.nacencomm.vn/'; E = 'MTCuu' },
    [ordered]@{ A = 'Công ty TNHH Hệ thống thông tin FPT'; B = '0104128565'; D = 'https://hoadon.ftg.vn/'; E = 'Mã hóa đơn' },
    [ordered]@{ A = 'Công ty Cổ phần công nghệ KIOTVIET'; B = '0104359717'; D = 'https://tracuuhoadon.kiotviet.vn/' },
    [ordered]@{ A = 'Công ty Cổ phần giải pháp First Trust'; B = '0104493085' },
    [ordered]@{ A = 'Công ty Cổ phần đầu tư và công nghệ idocNet'; B = '0104614692'; D = 'https://hoadontvan.com/TraCuu' },
    [ordered]@{ A = 'Công ty Cổ phần phát triển công nghệ ACMAN'; B = '0104908371'; D = 'https://hoadondientu.acman.vn/tra-cuu/hoa-don.html' },
    [ordered]@{ A = 'Công ty Cổ phần CyberLotus'; B = '0105232093'; D = 'https://tracuu.cyberbill.vn/'; E = 'MaTraCuu' },
    [ordered]@{ A = 'Công ty Cổ phần Megabiz Việt Nam'; B = '0105844836'; D = 'https://tracuu.vinvoice.vn/'; E = 'SearchKey' },
    [ordered]@{ A = 'Công ty Cổ phần hóa đơn điện tử New-Invoice'; B = '0105937449'; D = 'https://newinvoice.com.vn/tra-cuu/' },
    [ordered]@{ A = 'Công ty Cổ phần công nghệ ITT'; B = '0105958921'; D = 'https://tracuu.cloudinvoice.vn/' },
    [ordered]@{ A = 'Công ty Cổ phần Đầu tư công nghệ và thương mại Softdreams'; B = '0105987432'; D = 'http://" & mst & "hd.easyinvoice.com.vn'; E = 'Fkey'; F = 'http:// & mst & "hd.easyinvoice.com.vn"' },
    [ordered]@{ A = 'Công ty TNHH Hóa đơn điện tử M-INVOICE'; B = '0106026495'; D = 'https://tracuuhoadon.minvoice.com.vn/single/invoice'; E = 'Mã tra cứu' },
    [ordered]@{ A = 'Công ty TNHH Hóa đơn điện tử M-INVOICE'; B = '0106026495-001'; D = 'https://tracuuhoadon.minvoice.com.vn/single/invoice'; E = 'Mã tra cứu' },
    [ordered]@{ A = 'Công ty Cổ phần MONT-E'; B = '0106249501'; D = 'https://tracuuhoadon.minvoice.com.vn/single/invoice' },
    [ordered]@{ A = 'Công ty Cổ phần truyền số liệu Việt Nam'; B = '0106361479'; D = 'https://tracuu.ahoadon.com/' },
    [ordered]@{ A = 'Công ty Cổ phần dịch vụ T-VAN HILO'; B = '0106713804'; D = 'https://tracuuhddt78.hilo.com.vn/'; E = 'Fkey' },
    [ordered]@{ A = 'Công ty TNHH Giải pháp hóa đơn điện tử My – Invoice'; B = '0106820789'; D = 'https://tracuu.hoadondientuvn.info/' },
    [ordered]@{ A = 'Công ty Cổ phần VETC'; B = '0106858609'; D = 'https://tracuuhoadon.vetc.com.vn/?s' },
    [ordered]@{ A = 'Công ty Cổ phần ICORP'; B = '0106870211'; D = 'https://tracuu.vietinvoice.vn/' },
    [ordered]@{ A = 'CÔNG TY TNHH THU PHÍ TỰ ĐỘNG VETC'; B = '0107500414'; D = 'https://tracuuhoadon.vetc.com.vn/'; E = 'Mã tra cứu' },
    [ordered]@{ A = 'Công ty Cổ phần ATIS'; B = '0107732197' },
    [ordered]@{ A = 'Công ty Cổ phần Giải pháp phần mềm 3A'; B = '0108516079'; D = 'http://hddt.3asoft.vn/'; E = 'client_id' },
    [ordered]@{ A = 'Công ty Cổ phần My Software'; B = '0108971656'; D = 'https://tracuu.myinvoice.vn/' },
    [ordered]@{ A = 'Công ty cổ phần giao thông số việt nam'; B = '0109266456'; D = 'https://giaothongso.com.vn/tra-cuu-hoa-don-mtc/'; E = 'Mã hóa đơn' },
    [ordered]@{ A = 'Công ty Cổ phần Hóa đơn điện tử VININVOICE'; B = '0109282176'; D = 'https://tracuu.vininvoice.vn/' },
    [ordered]@{ A = 'Công ty Cổ phần công nghệ số và in đồ họa'; B = '0200638946'; D = 'https://oinvoice.vn/tracuu/' },
    [ordered]@{ A = 'Công ty Cổ phần Thiết bị điện - Điện tử Bách Khoa'; B = '0200784873'; D = 'https://hoadonbachkhoa.pmbk.vn/tra-cuu-hoa-don' },
    [ordered]@{ A = 'Công ty TNHH Tư vấn và Dịch vụ Home Casta'; B = '0201802839'; D = 'https://tracuu.homecasta.vn/' },
    [ordered]@{ A = 'Công ty Cổ phần phát triển và ứng dụng phần mềm Bách Khoa'; B = '0202029650'; D = 'https://hdbk.pmbk.vn/tra-cuu-hoa-don' },
    [ordered]@{ A = 'Công ty Cổ phần Tin học Lạc Việt'; B = '0301448733'; D = 'https://accnet.vn/hoa-don-dien-tu' },
    [ordered]@{ A = 'Công ty TNHH Giấy vi tính Liên Sơn'; B = '0301452923'; D = 'https://tracuu.lienson.vn/' },
    [ordered]@{ A = 'Công ty TNHH PA Việt Nam'; B = '0302431595'; D = 'https://tracuu.hoadon30s.vn' },
    [ordered]@{ A = 'Công ty Cổ phần Mắt Bão'; B = '0302712571'; D = 'https://matbao.in/tra-cuu-hoa-don/' },
    [ordered]@{ A = 'Công ty TNHH L.C.S'; B = '0302999571'; D = 'https://eip.lcssoft.com.vn/desktop/' },
    [ordered]@{ A = 'Công ty TNHH Kế toán và tư vấn V.L.C'; B = '0303211948' },
    [ordered]@{ A = 'Công ty Cổ phần công nghệ San Phú'; B = '0303430876'; D = 'http://trahoadon.vn/SearchOne' },
    [ordered]@{ A = 'Công ty TNHH Phần mềm và Tư vấn Kim Tự Tháp'; B = '0303549303' },
    [ordered]@{ A = 'Công ty TNHH Tin học Tia lửa Việt'; B = '0303609305'; D = 'https://ihoadondientu.com/Tra-cuu' },
    [ordered]@{ A = 'Công ty Cổ phần phần mềm Rosy'; B = '0305142231' },
    [ordered]@{ A = 'CÔNG TY CỔ PHẦN XĂNG DẦU DẦU KHÍ PHÚ THỌ'; B = '0305795054'; D = 'https://hoadon.pvoil.vn/Invoice/search'; E = 'Fkey' },
    [ordered]@{ A = 'Công ty TNHH Máy tính và truyền thông công nghệ kết nối'; B = '0306784030'; D = 'https://ehoadon.online/einvoice/lookup' },
    [ordered]@{ A = 'Công ty Cổ phần TS24'; B = '0309478306'; D = 'https://tracuu.xuathoadon.vn/' },
    [ordered]@{ A = 'Công ty Cổ phần chữ ký số VI NA'; B = '0309612872'; D = 'https://tracuuhd.smartsign.com.vn/' },
    [ordered]@{ A = 'Công ty Cổ phần Công nghệ UNIT'; B = '0309889835' },
    [ordered]@{ A = 'Công ty Cổ phần Chứng số An toàn'; B = '0310151055' },
    [ordered]@{ A = 'Công ty Cổ phần Dịch vụ Thương mại Việt Nam trực tuyến'; B = '0310151739'; D = 'https://news.yoinvoice.vn/search-invoice' },
    [ordered]@{ A = 'Công ty TNHH Dịch vụ phần mềm AVSE'; B = '0310768095'; D = 'http://hoadondientu.link/tracuutt78' },
    [ordered]@{ A = 'Công ty TNHH Phần mềm kế toán và dịch vụ thủ tục thuế Sài Gòn'; B = '0310926922'; D = 'https://invoice.ehcm.vn/' },
    [ordered]@{ A = 'Công ty TNHH Dịch vụ Trí Việt Luật'; B = '0311622035' },
    [ordered]@{ A = 'Công ty TNHH Phần mềm BRB'; B = '0311914694' },
    [ordered]@{ A = 'Công ty Cổ phần công nghệ VIETINFO'; B = '0311928954'; D = 'https://tracuuhoadon.vietinfo.tech/' },
    [ordered]@{ A = 'Công ty TNHH Giải pháp công nghệ Ngô Gia Phát'; B = '0311942758'; D = 'https://tracuuonline78.ngogiaphat.vn/Search' },
    [ordered]@{ A = 'Công ty TNHH NC9 Việt Nam'; B = '0312270160'; D = 'https://ameinvoice.vn/tra-cuu-hoa-don-dien-tu/' },
    [ordered]@{ A = 'Công ty TNHH Win Tech Solution'; B = '0312303803'; D = 'https://tracuu.wininvoice.vn/'; E = 'Mã tra cứu hóa đơn' },
    [ordered]@{ A = 'Công ty TNHH Công nghệ LCD Việt Nam'; B = '0312483391' },
    [ordered]@{ A = 'Công ty TNHH Ecount Việt Nam'; B = '0312575123' },
    [ordered]@{ A = 'Công ty TNHH Nhóm Mây'; B = '0312617990' },
    [ordered]@{ A = 'Công ty TNHH Công nghệ HT Sài Gòn'; B = '0312942260'; D = 'https://ihoadondientu.net/Tracuu.aspx' },
    [ordered]@{ A = 'Công ty TNHH MTV in Bến Thành'; B = '0312961577'; D = 'http://tracuuhoadon.benthanhinvoice.vn/' },
    [ordered]@{ A = 'Công ty Cổ phần Công nghệ TADU'; B = '0313253288' },
    [ordered]@{ A = 'Công ty TNHH Đầu tư Hòn Ngọc Việt'; B = '0313844107'; D = 'http://voice.hoadondientu.net.vn/tra-cuu' },
    [ordered]@{ A = 'Công ty Cổ phần Phát triển công nghệ Nguyễn Minh'; B = '0313906508'; D = 'https://nguyenminhvat.vn/hddt/sinv/sinv00101' },
    [ordered]@{ A = 'Công ty TNHH ZAMO'; B = '0313950909'; D = 'https://koffi.vn/outbound/lookup-invoice' },
    [ordered]@{ A = 'Công ty TNHH Soft Ware KK VAT'; B = '0313963672'; D = 'https://tracuuhoadon.kkvat.com.vn/' },
    [ordered]@{ A = 'Công ty TNHH Viễn thông Đông Sài Gòn'; B = '0314058603' },
    [ordered]@{ A = 'Công ty TNHH Thương Mại dịch vụ Online VI NA'; B = '0314185087'; D = 'https://hoadon.onlinevina.com.vn/invoice' },
    [ordered]@{ A = 'Công ty Cổ phần Minh Khang Group'; B = '0314209362'; D = 'https://hoadondientuvat.com/Tracuu.aspx' },
    [ordered]@{ A = 'Công ty TNHH Công nghệ Vĩnh Hy'; B = '0314743623'; D = 'https://ehoadondientu.com/Tra-cuu' },
    [ordered]@{ A = 'Công ty TNHH Phần mềm PVS'; B = '0315151651'; D = 'https://ei.pvssolution.com/' },
    [ordered]@{ A = 'Công ty TNHH Tư vấn thương mại Trí Việt Luật'; B = '0315191291'; D = 'https://hoadonsovn.evat.vn/' },
    [ordered]@{ A = 'Công ty TNHH Dịch vụ kế toán - Tư vấn thuế TTL'; B = '0315194912' },
    [ordered]@{ A = 'Công ty TNHH Hóa đơn điện tử TCT'; B = '0315298333'; D = 'https://tctinvoice.com/' },
    [ordered]@{ A = 'Công ty TNHH ACCONLINE.VN'; B = '0315467091'; D = 'https://www.acconline.vn/vn/tra-cuu-hoa-don.htm' },
    [ordered]@{ A = 'Công ty Cổ phần công nghệ hóa đơn điện tử HT'; B = '0315638251'; D = 'https://htinvoice.com.vn/TraCuu' },
    [ordered]@{ A = 'Công ty Cô phần Công nghệ Phát triển hóa đơn điện tử Việt Nam'; B = '0315983667' },
    [ordered]@{ A = 'Công ty TNHH Bizzi VietNam'; B = '0316114998' },
    [ordered]@{ A = 'Công ty Cổ phần Công nghệ BEE'; B = '0316636497' },
    [ordered]@{ A = 'Công ty TNHH Công nghệ và tư vấn Phương Nam'; B = '0316642395'; D = 'https://phuongnam.evat.vn/' },
    [ordered]@{ A = 'Công ty TNHH Tuần Châu'; B = '0400462489'; D = 'https://e-invoicetuanchau.com/Tra-cuu' },
    [ordered]@{ A = 'Công ty Cổ phần thương mại Visnam'; B = '0401486901'; D = 'https://tracuu.vin-hoadon.com/tracuuhoadon/tracuuxacthuc/tracuuhd'; E = 'Quanly_SoBaoMat' },
    [ordered]@{ B = '102519041'; D = 'https://ihoadon.vn/kiem-tra/?lang=vn' },
    [ordered]@{ A = 'Công ty TNHH WEBCASH Việt Nam'; B = '1201496252' },
    [ordered]@{ A = 'Công ty TNHH Minh Thư'; B = '3500456910'; D = 'https://hoadonminhthuvungtau.com/Tra-cuu' },
    [ordered]@{ A = 'Công ty TNHH MTV thương mại dịch vụ Trần Đình Tùng'; B = '3702037020'; D = 'https://trandinhtung.evat.vn/' },
    [ordered]@{ A = 'Công ty CP Tư vấn và Chuyển giao công nghệ Sơn Phát'; B = '4601328480'; D = 'https://tracuuhddt78.hilo.com.vn/' },
    [ordered]@{ A = 'CÔNG TY CỔ PHẦN DI CHUYỂN XANH VÀ THÔNG MINH GSM'; B = '0110269067'; D = 'https://gsm-einvoice.hilo.com.vn/'; E = 'Hilo-SearchKey' },
    [ordered]@{ A = 'CN - CÔNG TY CỔ PHẦN DI CHUYỂN XANH VÀ THÔNG MINH GSM'; B = '0110269067-002'; D = 'https://gsm-einvoice.hilo.com.vn/'; E = 'Hilo-SearchKey' },
    [ordered]@{ A = 'Mobifone invoice'; B = '0106026495-001'; D = 'https://tracuuhoadon.minvoice.com.vn/'; E = 'Mã tra cứu' },
    [ordered]@{ A = 'Mobifone invoice'; B = '0106026495'; D = 'https://tracuuhoadon.minvoice.com.vn/'; E = 'Mã tra cứu' }
)

function Get-TaiHoaDonDienTuLookupCell {
    param($Row, [string]$Column)
    if ($null -eq $Row) { return '' }
    $value = $null
    if ($Row -is [System.Collections.IDictionary]) {
        if (-not $Row.Contains($Column)) { return '' }
        $value = $Row[$Column]
    }
    else {
        $property = $Row.PSObject.Properties[$Column]
        if ($null -eq $property) { return '' }
        $value = $property.Value
    }
    if ($null -eq $value) { return '' }
    return [string]$value
}

function Get-TaiHoaDonDienTuLookupHeaders {
    $headers = @()
    foreach ($column in $script:TaiHoaDonDienTuLookupColumns) {
        $headers += (Get-TaiHoaDonDienTuLookupCell $script:TaiHoaDonDienTuLookupSheetData[0] $column)
    }
    return $headers
}

# Các dòng dữ liệu (không kể dòng tiêu đề) để ghi ra sheet LinkTraCuu.
function Get-TaiHoaDonDienTuLookupSheetRows {
    $rows = New-Object System.Collections.Generic.List[object]
    for ($index = 1; $index -lt $script:TaiHoaDonDienTuLookupSheetData.Count; $index++) {
        $rows.Add($script:TaiHoaDonDienTuLookupSheetData[$index])
    }
    return $rows.ToArray()
}

# Tên trường chứa mã tra cứu (cột E), tương đương dicTenCotTC.
$script:TaiHoaDonDienTuLookupFields = @(
    'TenDviQly'
    'Mã số bí mật'
    'reservationCode'
    'Fkey'
    'KeySearch'
    'TransactionID'
    'MaTC'
    'Mã hóa đơn'
    'Mã Kiểm tra'
    'Mã tra cứu'
    'Mã tra cứu hóa đơn'
    'MaTraCuu'
    'MTCuu'
    'Quanly_SoBaoMat'
    'SearchKey'
    'client_id'
    'Hilo-SearchKey'
)

function Test-TaiHoaDonDienTuLookupField {
    param([string]$FieldName)
    if ([string]::IsNullOrWhiteSpace($FieldName)) { return $false }
    $normalized = ($FieldName -replace '[^A-Za-z0-9]', '').ToLowerInvariant()
    foreach ($allowed in $script:TaiHoaDonDienTuLookupFields) {
        $allowedNormalized = ($allowed -replace '[^A-Za-z0-9]', '').ToLowerInvariant()
        if ($normalized -eq $allowedNormalized) { return $true }
    }
    return $false
}

# Nhãn trạng thái hóa đơn (I2:I8); chỉ số tthai ánh xạ thẳng, 0 = 'Tất cả'.
function Get-TaiHoaDonDienTuStatusLabels {
    $labels = @()
    foreach ($row in (Get-TaiHoaDonDienTuLookupSheetRows)) {
        $value = Get-TaiHoaDonDienTuLookupCell $row 'I'
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        $labels += $value
        if ($labels.Count -eq 7) { break }
    }
    return $labels
}

# Nhãn kết quả kiểm tra dùng khi xuất Excel (O2:O11 cho cả MUA và BAN, đúng
# như VBA gán arrKQKTHoaDon trước khi ghi cột 53).  Ô đầu là 'Tất cả' nên
# ttxly N được tra bằng chỉ số N + 1.
function Get-TaiHoaDonDienTuValidationLabels {
    $labels = @()
    foreach ($row in (Get-TaiHoaDonDienTuLookupSheetRows)) {
        $value = Get-TaiHoaDonDienTuLookupCell $row 'O'
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        $labels += $value
        if ($labels.Count -eq 10) { break }
    }
    return $labels
}

# Nhãn kết quả kiểm tra của bộ lọc tab MUA (K2:L5 trong VBA).
function Get-TaiHoaDonDienTuPurchaseValidationLabels {
    $labels = @()
    foreach ($row in (Get-TaiHoaDonDienTuLookupSheetRows)) {
        $value = Get-TaiHoaDonDienTuLookupCell $row 'L'
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        $labels += $value
        if ($labels.Count -eq 4) { break }
    }
    return $labels
}

# MST nhà cung cấp -> link tra cứu, và MST người bán -> link riêng.
# Dựng lại từ bảng dữ liệu để dùng chung cho bảng gốc và bảng do người dùng
# bổ sung; dòng cuối thắng, đúng như VBA gán đè vào dictionary.
function Sync-TaiHoaDonDienTuLookupIndex {
    $script:TaiHoaDonDienTuLookupLinks = @{}
    $script:TaiHoaDonDienTuSellerLookupLinks = @{}
    foreach ($lookupRow in (Get-TaiHoaDonDienTuLookupSheetRows)) {
        $provider = (Get-TaiHoaDonDienTuLookupCell $lookupRow 'B')
        $seller = (Get-TaiHoaDonDienTuLookupCell $lookupRow 'C')
        $link = (Get-TaiHoaDonDienTuLookupCell $lookupRow 'D')
        if ([string]::IsNullOrWhiteSpace($link)) { continue }
        if (-not [string]::IsNullOrWhiteSpace($provider)) { $script:TaiHoaDonDienTuLookupLinks[$provider] = $link }
        if (-not [string]::IsNullOrWhiteSpace($seller)) { $script:TaiHoaDonDienTuSellerLookupLinks[$seller] = $link }
    }
    $fieldNames = New-Object System.Collections.Generic.List[string]
    foreach ($lookupRow in (Get-TaiHoaDonDienTuLookupSheetRows)) {
        $field = (Get-TaiHoaDonDienTuLookupCell $lookupRow 'E')
        if ([string]::IsNullOrWhiteSpace($field)) { continue }
        if (-not (Test-TaiHoaDonDienTuLookupField $field)) { $fieldNames.Add($field) }
    }
    if ($fieldNames.Count -gt 0) { $script:TaiHoaDonDienTuLookupFields += $fieldNames.ToArray() }
}

Sync-TaiHoaDonDienTuLookupIndex

# Tương đương dicLink.item(msttcgp).
function Get-TaiHoaDonDienTuLookupLink {
    param([string]$ProviderTaxCode)
    if ([string]::IsNullOrWhiteSpace($ProviderTaxCode)) { return '' }
    $key = $ProviderTaxCode.Trim()
    if ($script:TaiHoaDonDienTuLookupLinks.ContainsKey($key)) {
        return [string]$script:TaiHoaDonDienTuLookupLinks[$key]
    }
    return ''
}

# Link riêng theo MST người bán (cột C): hóa đơn của Viettel phát ra cho
# Petroimex hay Vingroup có trang tra cứu riêng, ưu tiên link chung.
function Get-TaiHoaDonDienTuSellerLookupLink {
    param([string]$SellerTaxCode)
    if ([string]::IsNullOrWhiteSpace($SellerTaxCode)) { return '' }
    $key = $SellerTaxCode.Trim()
    if ($script:TaiHoaDonDienTuSellerLookupLinks.ContainsKey($key)) {
        return [string]$script:TaiHoaDonDienTuSellerLookupLinks[$key]
    }
    return ''
}

# --- Đọc bảng LinkTraCuu từ workbook của người dùng -------------------------
# Cho phép người dùng thêm/sửa link trong sheet LinkTraCuu rồi trỏ
# LOOKUP_TABLE_XLSX vào file đó; các dòng đó được nối vào bảng gốc (dòng
# cuối thắng) nên link trong cột 55/56 lấy theo.  Đọc trực tiếp gói OOXML
# bằng .NET nên chạy được trên Windows, macOS và Linux.

function Get-TaiHoaDonDienTuZipEntryText {
    param($Archive, [string]$Name)
    $entry = $null
    foreach ($candidate in $Archive.Entries) {
        if (([string]$candidate.FullName) -eq $Name) { $entry = $candidate; break }
    }
    if ($null -eq $entry) { return $null }
    $stream = $entry.Open()
    $reader = New-Object IO.StreamReader($stream, [Text.Encoding]::UTF8)
    try { return $reader.ReadToEnd() }
    finally { $reader.Dispose(); $stream.Dispose() }
}

function Get-TaiHoaDonDienTuSharedStrings {
    param($Archive)
    $text = Get-TaiHoaDonDienTuZipEntryText $Archive 'xl/sharedStrings.xml'
    $items = New-Object System.Collections.Generic.List[string]
    if ($null -eq $text) { return $items.ToArray() }
    $document = [xml]$text
    $manager = New-Object Xml.XmlNamespaceManager($document.NameTable)
    $manager.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    foreach ($node in $document.SelectNodes('/x:sst/x:si', $manager)) {
        $builder = New-Object Text.StringBuilder
        foreach ($textNode in $node.SelectNodes('.//x:t', $manager)) { [void]$builder.Append($textNode.InnerText) }
        $items.Add($builder.ToString())
    }
    return $items.ToArray()
}

# Đọc các cột A..F của sheet LinkTraCuu trong một workbook .xlsx.
function Read-TaiHoaDonDienTuLookupSheet {
    param([Parameter(Mandatory = $true)][string]$Path, [string]$SheetName = 'LinkTraCuu')
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Không tìm thấy file bảng tra cứu: $Path"
    }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $workbookText = Get-TaiHoaDonDienTuZipEntryText $archive 'xl/workbook.xml'
        if ($null -eq $workbookText) { throw "File không phải workbook .xlsx hợp lệ: $Path" }
        $workbook = [xml]$workbookText
        $workbookManager = New-Object Xml.XmlNamespaceManager($workbook.NameTable)
        $workbookManager.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        $workbookManager.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')

        $relationshipId = ''
        foreach ($sheet in $workbook.SelectNodes('/x:workbook/x:sheets/x:sheet', $workbookManager)) {
            if (([string]$sheet.GetAttribute('name')) -eq $SheetName) {
                $relationshipId = [string]$sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
                break
            }
        }
        if ([string]::IsNullOrWhiteSpace($relationshipId)) {
            throw "File $Path không có sheet '$SheetName'."
        }

        $target = ''
        $workbookRelsText = Get-TaiHoaDonDienTuZipEntryText $archive 'xl/_rels/workbook.xml.rels'
        if ($null -eq $workbookRelsText) { throw "File $Path thiếu quan hệ của workbook." }
        $workbookRels = [xml]$workbookRelsText
        foreach ($relationship in $workbookRels.DocumentElement.ChildNodes) {
            if (([string]$relationship.GetAttribute('Id')) -eq $relationshipId) {
                $target = [string]$relationship.GetAttribute('Target')
                break
            }
        }
        if ([string]::IsNullOrWhiteSpace($target)) { throw "File $Path không xác định được sheet '$SheetName'." }
        $target = $target.TrimStart('/')
        if (-not $target.StartsWith('xl/')) { $target = 'xl/' + $target }

        $sheetText = Get-TaiHoaDonDienTuZipEntryText $archive $target
        if ($null -eq $sheetText) { throw "File $Path không có dữ liệu sheet '$SheetName'." }
        $sharedStrings = Get-TaiHoaDonDienTuSharedStrings $archive
        $sheetDocument = [xml]$sheetText
        $sheetManager = New-Object Xml.XmlNamespaceManager($sheetDocument.NameTable)
        $sheetManager.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

        $rows = New-Object System.Collections.Generic.List[object]
        foreach ($rowNode in $sheetDocument.SelectNodes('/x:worksheet/x:sheetData/x:row', $sheetManager)) {
            $values = [ordered]@{}
            foreach ($cell in $rowNode.SelectNodes('x:c', $sheetManager)) {
                $reference = [string]$cell.GetAttribute('r')
                if ($reference -notmatch '^([A-Z]+)') { continue }
                $column = $Matches[1]
                if ($script:TaiHoaDonDienTuLookupColumns -notcontains $column) { continue }
                $text = ''
                $type = [string]$cell.GetAttribute('t')
                if ($type -eq 'inlineStr') {
                    foreach ($textNode in $cell.SelectNodes('x:is//x:t', $sheetManager)) { $text += $textNode.InnerText }
                }
                else {
                    $valueNode = $cell.SelectSingleNode('x:v', $sheetManager)
                    $text = if ($null -eq $valueNode) { '' } else { [string]$valueNode.InnerText }
                    if ($type -eq 's' -and -not [string]::IsNullOrWhiteSpace($text)) {
                        $index = 0
                        if ([int]::TryParse($text, [ref]$index) -and $index -ge 0 -and $index -lt $sharedStrings.Count) {
                            $text = [string]$sharedStrings[$index]
                        }
                    }
                }
                $values[$column] = $text.Trim()
            }
            $hasValue = $false
            foreach ($entry in $values.GetEnumerator()) {
                if (-not [string]::IsNullOrWhiteSpace([string]$entry.Value)) { $hasValue = $true; break }
            }
            if ($hasValue) { $rows.Add($values) }
        }
        return $rows.ToArray()
    }
    finally { $archive.Dispose() }
}

# Nối bảng của người dùng vào bảng gốc và dựng lại chỉ mục tra cứu.  Dòng có
# cùng MST (hoặc cùng MST người bán) thì thay thế dòng cũ, nên trỏ
# LOOKUP_TABLE_XLSX vào workbook của lần chạy trước cũng không nhân bản bảng.
function Import-TaiHoaDonDienTuLookupTable {
    param([Parameter(Mandatory = $true)][string]$Path, [string]$SheetName = 'LinkTraCuu')
    $rows = @(Read-TaiHoaDonDienTuLookupSheet -Path $Path -SheetName $SheetName)
    $dataRows = New-Object System.Collections.Generic.List[object]
    foreach ($row in $rows) {
        $text = [string](Get-TaiHoaDonDienTuLookupCell $row 'A')
        $provider = [string](Get-TaiHoaDonDienTuLookupCell $row 'B')
        $link = [string](Get-TaiHoaDonDienTuLookupCell $row 'D')
        if ($provider -eq 'MST' -or ($text -eq 'Tên tổ chức' -and $provider -eq '')) { continue }
        if ([string]::IsNullOrWhiteSpace($provider) -and [string]::IsNullOrWhiteSpace($link)) { continue }
        $dataRows.Add($row)
    }
    if ($dataRows.Count -eq 0) {
        throw "Sheet '$SheetName' trong $Path không có dòng MST/link nào để dùng."
    }
    foreach ($row in $dataRows) {
        $provider = [string](Get-TaiHoaDonDienTuLookupCell $row 'B')
        $seller = [string](Get-TaiHoaDonDienTuLookupCell $row 'C')
        $replaced = $false
        # Khoá của một dòng là cặp (MST nhà cung cấp, MST người bán): cùng MST
        # nhà cung cấp nhưng khác MST người bán là hai dòng khác nhau, đúng như
        # sheet LinkTraCuu gốc.
        for ($index = 1; $index -lt $script:TaiHoaDonDienTuLookupSheetData.Count; $index++) {
            $current = $script:TaiHoaDonDienTuLookupSheetData[$index]
            if ((Get-TaiHoaDonDienTuLookupCell $current 'B') -ne $provider) { continue }
            if ((Get-TaiHoaDonDienTuLookupCell $current 'C') -ne $seller) { continue }
            $script:TaiHoaDonDienTuLookupSheetData[$index] = $row
            $replaced = $true
            break
        }
        if (-not $replaced) { $script:TaiHoaDonDienTuLookupSheetData += $row }
    }
    Sync-TaiHoaDonDienTuLookupIndex
    return $dataRows.Count
}
