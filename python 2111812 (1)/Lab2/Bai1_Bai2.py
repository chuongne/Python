from datetime import datetime

class SinhVien:
    truong = "Dai hoc Da Lat"

    def __init__(self, maSo: int, hoTen: str, ngaySinh: datetime) -> None:
        self._maSo=maSo
        self._hoTen=hoTen
        self._ngaySinh=ngaySinh
    
    @property
    def maSo(self):
        return self._maSo
    
    @maSo.setter
    def maSo(self, maso):
        if self.laMaSoHopLe(maso):
            self._maSo=maso
    
    @staticmethod
    def laMaSoHopLe(maso: int):
        return len(str(maso))==7
    
    @classmethod
    def doiTenTruong(self, tenmoi):
        self.truong=tenmoi
    
    def __str__(self) -> str:
        return f"{self._maSo}\t{self._hoTen}\t{self._ngaySinh}"

    def xuat(self):
        print(f"{self._maSo}\t{self._hoTen}\t{self._ngaySinh}")

class DanhSachSv:
    def __init__(self) -> None:
        self.dssv = []

    def themSinhVien(self, sv: SinhVien):
        self.dssv.append(sv)

    def xuat(self):
        for sv in self.dssv:
            print(sv)

    def timSvTheoMSSV(self, maSo: int):
        return [sv for sv in self.dssv if sv.maSo == maSo]
    
    def timVTSvTheoMSSV(self, maSo:int):
        for i in range(len(self.dssv)):
            if self.dssv[i].maSo == maSo:
                return i
            return -1
        
    def xoaSvTheoMSSV(self, maSo: int)->bool:
        vt = self.timSvTheoMSSV(maSo)
        if vt != -1:
            del self.dssv[vt]
            return True
        else:
            return False
    
    def timSvTheoTen(self, ten:str):
        return [sv for sv in self.dssv if (sv.hoTen.split(" ")[-1].upper()== ten.upper())]

    def timSvSinhTruocNgay(self, ngay: datetime):
        return [sv for sv in self.dssv if sv.ngaySinh < ngay]

sv1 = SinhVien(2111839,"Tran Van Nam",datetime.strptime("7/7/2003",'%m/%d/%Y'))
sv2 = SinhVien(2111838,"Le tan Dan Huy",datetime.strptime("7/7/2003",'%m/%d/%Y'))
sv3 = SinhVien(2111837,"Nguyen Nhat Truong",datetime.strptime("7/7/2003",'%m/%d/%Y'))
sv4 = SinhVien(2111836,"Dinh Le Quang Huy",datetime.strptime("7/7/2003",'%m/%d/%Y'))
sv5 = SinhVien(2111835,"Pham Anh Quan",datetime.strptime("7/7/2003",'%m/%d/%Y'))

dssv = DanhSachSv()
dssv.themSinhVien(sv1)
dssv.themSinhVien(sv2)
dssv.themSinhVien(sv3)
dssv.themSinhVien(sv4)
dssv.themSinhVien(sv5)
#dsCTK45.xuat()

#Tim theo ma sinh vien
a = int(input('Nhập vào mã số muốn tìm kiếm: '))
kq = dssv.timSvTheoMSSV(a)
if dssv.timVTSvTheoMSSV(a) != -1:
    print(f'Đã tìm thấy sinh viên có mã số {a}')
    for sv in kq:
        sv.xuat()
else:
    print(f'Không tìm thấy sinh viên có mã số {a}')

b = int(input('Nhập vào mã số muốn xóa: '))

kq = dssv.xoaSvTheoMSSV(b)
print(f'\nĐã xóa {b} khỏi danh sách\n')
dssv.xuat()

c = str(input('Nhập vào tên sinh viên cần tìm: '))
kq = dssv.timSvTheoTen(c)
print(f'\nĐã tìm thấy tên {c}')
for sv in kq:
    sv.xuat()

ngay = datetime.strptime('23/04/2001', '%d/%m/%Y')
kqNgay = dssv.timSvSinhTruocNgay(ngay)

## Bai 2
f = open("demofile.txt", "r")
print(f.read())

def sapXepSvTheoHoTen(self):
    self.dssv.soft(key = lambda x: x.hoTen, reversed = False)
sapXepSvTheoHoTen()
for sv in kqNgay:
    sv.xuat()