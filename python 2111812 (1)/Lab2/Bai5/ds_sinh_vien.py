from datetime import datetime
from sinh_vien import SinhVien
from sinh_vien_chinh_quy import SinhVienChinhQuy 
from sv_phi_chinh_quy import SinhVienPhiCQ
class DanhSachSv:
    def __init__(self) -> None: 
        self.dssv = []
    def themSV(self, sv: SinhVien): 
        self.dssv.append(sv)
    def xuat(self):
        for sv in self.dssv: 
            print(sv)
    def timSVTheoMs (self, ms: str):
        for i in range(len(self.dssv)): 
            if self.dssv[i].mssv == ms:
                return i
        else:
            return -1        
    def timSvTheoLoai (self, loai: str): 
        if loai == "chinhquy":
            return [sv for sv in self.dssv if isinstance (sv, SinhVienChinhQuy)] 
        return [sv for sv in self.dssv if isinstance (sv, SinhVienPhiCQ)]
    
    
    #tìm sinh viên có điểm rèn luyện từ 80 trở lên
    @staticmethod
    def kiemTraDiem(sv: SinhVien):
        if sv.diemRL >= 80:
            return True
        return False
    def timSvCDRLT80(self):
        return [sv for sv in self.dssv if self.kiemTraDiem(sv)]
     #tìm sinh viên có trình độ cao đẳng sinh trước ngày 15/8/1999
    @staticmethod
    def kiemTraNgaySinh(sv: SinhVien):
        if sv.TrinhDo.lower() == "Cao đẳng".lower() and sv.ngaySinh < datetime.strptime("15/8/1999", "%d/%m/%Y"):
            return True
        return False
    def timSvTDCDSTNgay(self):
        return [sv for sv in self.dssv if self.kiemTraNgaySinh(sv)]
    
    

    