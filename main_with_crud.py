# Mengimpor library yang dibutuhkan
import pandas as pd
from datetime import datetime
import numpy as np

# === Blueprint untuk Mengelola File Excel ===
class excelManager:
    def __init__(self, path: str, sheetName: str = "Sheet1", primaryKey="ID"):
        self.path = path
        self.sheetName = sheetName
        self.df = pd.read_excel(path)
        self.primaryKey = None

        # Deteksi kolom primary key yang cocok
        for i in self.df.columns:
            if (i.strip().lower() == primaryKey.strip().lower()):
                self.primaryKey = i

    # --- INSERT DATA (CREATE) ---
    def insertData(self, newData: dict):
        new_row = {}
        checkIfExist = self.getData(self.primaryKey, newData[self.primaryKey])

        if (checkIfExist):
            return f"Data dengan {self.primaryKey} {newData[self.primaryKey]} sudah ada."

        # Pastikan semua kolom cocok
        for newValue in newData:
            for col in self.df.columns:
                if (str(newValue).lower() == str(col).lower()):
                    new_row.update({col: newData[newValue]})
                    break

        # Tambahkan ke DataFrame
        self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
        self.saveChange()
        return f"Data dengan {self.primaryKey} {newData[self.primaryKey]} berhasil ditambahkan."

    # --- READ / GET DATA ---
    def getData(self, colName: str, data) -> dict:
        data = str(data)
        for row in self.df.index:
            temp = {}
            for col in self.df.columns:
                temp.update({col: self.df.at[row, col]})
            temp.update({"row": row})
            if (str(temp[colName]).lower().strip() == data):
                return {"result": temp, "row": row}

    # --- UPDATE DATA (EDIT) ---
    def editData(self, keyValue, updatedData: dict):
        found = self.getData(self.primaryKey, keyValue)
        if not found:
            return f"Data dengan {self.primaryKey} {keyValue} tidak ditemukan."

        rowIndex = found["row"]
        for col in updatedData:
            if col in self.df.columns:
                self.df.at[rowIndex, col] = updatedData[col]

        self.saveChange()
        return f"Data dengan {self.primaryKey} {keyValue} berhasil diperbarui."

    # --- DELETE DATA ---
    def deleteData(self, keyValue):
        found = self.getData(self.primaryKey, keyValue)
        if not found:
            return f"Data dengan {self.primaryKey} {keyValue} tidak ditemukan."

        rowIndex = found["row"]
        self.df = self.df.drop(index=rowIndex).reset_index(drop=True)
        self.saveChange()
        return f"Data dengan {self.primaryKey} {keyValue} berhasil dihapus."

    # --- SAVE CHANGES ---
    def saveChange(self):
        self.df.to_excel(self.path, sheet_name=self.sheetName, index=False)


# === MEMBUAT OBJEK PENGELOLA EXCEL ===
dataBarang = excelManager("materiVideo/dataBarangMinimarket.xlsx", primaryKey="ID")
dataPenjualan = excelManager("materiVideo/dataPenjualanMinimarket.xlsx", primaryKey="IDPejualan")


# === FUNGSI UNTUK PROSES PENJUALAN ===
def jual(idBarang, jumlahBarang):
    dataFound = dataBarang.getData("ID", idBarang)
    if not dataFound:
        return f"ID: {idBarang} tidak ditemukan."

    data = dataFound["result"]

    if (int(data["Stok"]) - jumlahBarang < 0):
        return "Jumlah barang melebihi stok."

    # Catat penjualan
    dataPenjualan.insertData({
        "IDPejualan": str(datetime.now().strftime("%Y%m%d%H%M%S%f")),
        "ID": data["ID"],
        "Kategori": data["Kategori"],
        "Harga": data["Harga"],
        "Waktu": datetime.now(),
        "Jumlah barang": jumlahBarang,
        "Total": float(data["Harga"]) * jumlahBarang
    })

    # Kurangi stok barang
    dataBarang.editData(data["ID"], {
        "Stok": int(data["Stok"]) - jumlahBarang
    })

    return f"Penjualan ID {idBarang} berhasil dicatat dan stok diperbarui."


def restock():
    pass

def barangPalingLaku():
    pass


# === CONTOH PENGGUNAAN ===
# print(dataBarang.insertData({
#     "ID": 1234567890,
#     "Nama": "Sabun Cuci",
#     "Perusahaan Asal": "PT Bersih Selalu",
#     "Kategori": "Kebutuhan Rumah Tangga",
#     "Harga": 10000,
#     "Stok": 50
# }))

# print(dataBarang.editData(1234567890, {"Stok": 80}))
# print(dataBarang.deleteData(1234567890))
# print(jual(8885193814391, 10))
