import math 
class PhanSo:
    def __init__(self, tu=1, mau=1):
        self.tu = tu
        self.mau = mau
    def __str__(self) -> str:
        return f"{self.tu}/{self.mau}"
    def rutGon(self):
        #ucln = self.UCLN(self.tu, self.mau)
        ucln = math.gcd(self.tu,self.mau)
        self.tu = self.tu // ucln
        self.mau = self.mau // ucln
    def UCLN(self, a, b):
        if b == 0:
            return a
        return self.UCLN(b, a % b)
    def __add__(self, other):
        ps = PhanSo()
        if not isinstance(other,PhanSo):
            other = PhanSo(other)
        ps.tu = self.tu * other.mau + self.mau *other.tu
        #print(ps.tu_so)
        ps.mau = self.mau * other.mau
        #return ps
        ps.rutGon()
        return ps
    def __sub__(self, other):
        ps = PhanSo()
        if not isinstance(other,PhanSo):
            other = PhanSo(other)
        ps.tu = self.tu * other.mau - self.mau *other.tu
        #print(ps.tu_so)
        ps.mau = self.mau * other.mau
        ps.rutGon()
        return ps
    def __mul__(self, other):
        ps = PhanSo()
        if not isinstance(other,PhanSo):
            other = PhanSo(other)
        ps.tu = self.tu * other.tu
        #print(ps.tu_so)
        ps.mau = self.mau * other.mau
        ps.rutGon()
        return ps
    def __truediv__(self, other):
        ps = PhanSo()
        if not isinstance(other,PhanSo):
            other = PhanSo(other)
        ps.tu = self.tu * other.mau
        #print(ps.tu_so)
        ps.mau = self.mau * other.tu
        ps.rutGon()
        return ps
class DanhSachPhanSo:
    def __init__(self, danh_sach=[]):
        self.danh_sach = danh_sach
    def __str__(self) -> str:
        return f"{self.tu}/{self.mau}"
    def demSoPhanSoAm(self):
        count = 0
        for phan_so in self.danh_sach:
            if phan_so.tu * phan_so.mau < 0:
                count += 1
        return count
    # def xuat()
    def timPhanSoDuongNhoNhat(self):
        phan_so_duong_min = None
        for phan_so in self.danh_sach:
            if phan_so.tu * phan_so.mau > 0:
                if phan_so_duong_min is None or phan_so.tu * phan_so.mau < phan_so_duong_min.tu * phan_so_duong_min.mau:
                    phan_so_duong_min = phan_so
        return phan_so_duong_min
    def timViTriCuaPhanSo(self, phan_so_x):
        vi_tri = []
        for i in range(len(self.danh_sach)):
            if self.danh_sach[i].tu == phan_so_x.tu and self.danh_sach[i].mau == phan_so_x.mau:
                vi_tri.append(i)
        return vi_tri
    def tongPhanSoAm(self):
        tong = PhanSo()
        for phan_so in self.danh_sach:
            if phan_so.tu * phan_so.mau < 0:
                tong += phan_so
        return tong
    # def timDsPhanSo(self):
    #     ds_kq = DanhSachPhanSo()
    #     for ps in self.danh_sach:
            
    def xoaPhanSo(self, phan_so_x):
        self.danh_sach = [phan_so for phan_so in self.danh_sach if phan_so.tu != phan_so_x.tu or phan_so.mau != phan_so_x.mau]
    def xoaPhanSoTu(self, tu):
        self.danh_sach = [phan_so for phan_so in self.danh_sach if phan_so.tu != tu]
    def sapXepTang(self):
        self.danh_sach.sort(key=lambda x: x.tu / x.mau)
    def sapXepGiam(self):
        self.danh_sach.sort(key=lambda x: x.tu / x.mau, reverse=True)
    def sapXepTangTheoMau(self):
        self.danh_sach.sort(key=lambda x: x.mau / x.tu) 
    def sapXepTangTheoTu(self):
        self.danh_sach.sort(key=lambda x: x.tu / x.mau) 
    def sapXepGiamTheoMau(self):
        self.danh_sach.sort(key=lambda x: x.mau / x.tu, reverse=True) 
    def sapXepGiamTheoTu(self):
        self.danh_sach.sort(key=lambda x: x.tu / x.mau, reverse=True)   

a = PhanSo()
a.tu = 2
a.mau = 3
b = PhanSo(3,5)
c = PhanSo(-1, 2)
d = PhanSo(4, 6)
danh_sach_phan_so = DanhSachPhanSo([a, b, c, d])
# print(f"{a} + {b} = {a+b}")
# print(f"{a} - {b} = {a-b}")
# print(f"{a} * {b} = {a*b}")
# print(f"{a} / {b} = {a/b}")
# print(f"Dem so phan so am trong mang: {danh_sach_phan_so.demSoPhanSoAm()}")
# print(f"Tim phan so duong nho nhat: {danh_sach_phan_so.timPhanSoDuongNhoNhat()}")
#print(f"Tim tat ca vi tri cua phan so x trong mang: {danh_sach_phan_so.timViTriCuaPhanSo(b)}")
#print(f"Tong tat ca cac phan so am trong mang: {danh_sach_phan_so.tongPhanSoAm()}")
## print(f"Xoa phan so x trong mang: {danh_sach_phan_so.xoaPhanSo(b)}")
## print(danh_sach_phan_so.danh_sach)
## print(f"Xoa tat ca phan so co tu la x: {danh_sach_phan_so.xoaPhanSoTu(a.tu)}")
## print(danh_sach_phan_so.danh_sach)
danh_sach_phan_so.sapXepTang()
print(f"Sap xep phan so theo chieu tang: {danh_sach_phan_so.danh_sach}")

# print(f"sap xep phan so theo chieu giam: {danh_sach_phan_so.sapXepGiam()}")
# print(f"sap xep phan so theo chieu tang theo mau: {danh_sach_phan_so.sapXepTangTheoMau()}")
# print(f"sap xep phan so theo chieu tang theo tu: {danh_sach_phan_so.sapXepTangTheoTu()}")
# print(f"sap xep phan so theo chieu giam theo tu: {danh_sach_phan_so.sapXepGiamTheoTu()}")
# print(f"sap xep phan so theo chieu giam theo  mau: {danh_sach_phan_so.sapXepGiamTheoMau()}")

