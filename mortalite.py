import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path
import warnings

warnings.filterwarnings('ignore')


class MortalityTableGenerator:
    def __init__(self):
        self.data = None
        self.mortality_table = None

    def read_excel_data(self, file_path, sheet_name=0):
        """
        Excel dosyasından mortalite verilerini okur

        Parameters:
        file_path (str): Excel dosyasının yolu
        sheet_name (str/int): Sayfa adı veya indeksi
        """
        try:
            # Excel dosyasını oku
            self.data = pd.read_excel(file_path, sheet_name=sheet_name)

            # Sütun isimlerini standartlaştır
            self.data.columns = self.data.columns.str.lower().str.strip()

            print(f"✅ Excel dosyası başarıyla okundu: {len(self.data)} satır")
            print(f"📊 Mevcut sütunlar: {list(self.data.columns)}")

            return self.data

        except Exception as e:
            print(f"❌ Hata: Excel dosyası okunamadı - {str(e)}")
            return None

    def validate_columns(self):
        """Gerekli sütunların varlığını kontrol eder"""

        # Alternatif sütun isimleri
        column_mapping = {
            'yaş': ['yas', 'age', 'yaş', 'x'],
            'lx': ['lx', 'l(x)', 'survivors'],
            'qx': ['qx', 'q(x)', 'mortality_rate'],
            'dx': ['dx', 'd(x)', 'deaths']
        }

        mapped_columns = {}

        for standard_col, alternatives in column_mapping.items():
            found = False
            for alt in alternatives:
                if alt in self.data.columns:
                    mapped_columns[standard_col] = alt
                    found = True
                    break

            # Yaş sütunu mutlaka gerekli
            if standard_col == 'yaş' and not found:
                print(f"❌ Yaş sütunu bulunamadı! Aranacak isimler: {alternatives}")
                return False

            # qx sütunu da gerekli (en azından)
            if standard_col == 'qx' and not found:
                print(f"❌ Ölüm oranı (qx) sütunu bulunamadı! Aranacak isimler: {alternatives}")
                return False

        # Sütunları yeniden adlandır
        rename_dict = {v: k for k, v in mapped_columns.items() if k in mapped_columns}
        self.data.rename(columns=rename_dict, inplace=True)

        print("✅ Sütunlar bulundu ve standartlaştırıldı")
        print(f"📋 Kullanılacak sütunlar: {list(mapped_columns.keys())}")
        return True

    def create_mortality_table(self, radix=100000):
        """
        Mortalite tablosunu oluşturur

        Parameters:
        radix (int): Başlangıç yaşayan sayısı (genellikle 100,000)
        """
        if self.data is None:
            print("❌ Önce Excel verisi okunmalı!")
            return None

        if not self.validate_columns():
            return None

        # Veriyi kopyala
        df = self.data.copy()

        # Yaş sütununu sırala
        df = df.sort_values('yaş').reset_index(drop=True)

        # Eksik sütunları oluştur
        if 'lx' not in df.columns:
            df['lx'] = np.nan
        if 'dx' not in df.columns:
            df['dx'] = np.nan

        # Eksik hesaplamaları tamamla
        df = self._complete_calculations(df, radix)

        # Ek istatistiksel hesaplamalar
        df = self._add_life_table_functions(df)

        self.mortality_table = df

        print("✅ Mortalite tablosu başarıyla oluşturuldu!")
        return df

    def _complete_calculations(self, df, radix):
        """Eksik lx, qx, dx değerlerini hesaplar"""

        # İlk lx değerini radix olarak ayarla (eğer yoksa)
        if pd.isna(df.iloc[0]['lx']):
            df.loc[0, 'lx'] = radix

        for i in range(len(df)):
            # lx hesaplama (önceki lx - önceki dx)
            if i > 0 and pd.isna(df.iloc[i]['lx']):
                if not pd.isna(df.iloc[i - 1]['lx']) and not pd.isna(df.iloc[i - 1]['dx']):
                    df.loc[i, 'lx'] = max(0, df.iloc[i - 1]['lx'] - df.iloc[i - 1]['dx'])

            # dx hesaplama (lx * qx)
            if pd.isna(df.iloc[i]['dx']) and not pd.isna(df.iloc[i]['qx']) and not pd.isna(df.iloc[i]['lx']):
                df.loc[i, 'dx'] = df.iloc[i]['lx'] * df.iloc[i]['qx']

            # qx hesaplama (dx / lx) - sadece eksikse
            if pd.isna(df.iloc[i]['qx']) and not pd.isna(df.iloc[i]['dx']) and not pd.isna(df.iloc[i]['lx']):
                if df.iloc[i]['lx'] > 0:
                    df.loc[i, 'qx'] = min(1.0, df.iloc[i]['dx'] / df.iloc[i]['lx'])
                else:
                    df.loc[i, 'qx'] = 0.0

        # Kalan eksik lx değerlerini hesapla
        for i in range(1, len(df)):
            if pd.isna(df.iloc[i]['lx']) and not pd.isna(df.iloc[i - 1]['lx']) and not pd.isna(df.iloc[i - 1]['qx']):
                df.loc[i, 'lx'] = max(0, df.iloc[i - 1]['lx'] * (1 - df.iloc[i - 1]['qx']))

        return df

    def _add_life_table_functions(self, df):
        """Ek yaşam tablosu fonksiyonlarını hesaplar"""
        try:
            # px (yaşama olasılığı) = 1 - qx
            df['px'] = 1 - df['qx']

            # Lx (yaş aralığında yaşanan yıl sayısı)
            df['Lx'] = df['lx'] - (df['dx'] / 2)

            # Tx (x yaşından sonra yaşanan toplam yıl)
            df['Tx'] = df['Lx'][::-1].cumsum()[::-1]

            # ex (yaşam beklentisi)
            df['ex'] = np.where(df['lx'] > 0, df['Tx'] / df['lx'], 0)

            # NaN ve inf değerlerini temizle
            df = df.replace([np.inf, -np.inf], np.nan)
            df = df.fillna(0)

        except Exception as e:
            print(f"⚠️ Ek hesaplamalarda hata: {str(e)}")

        return df

    def display_table(self, head=10):
        """Mortalite tablosunu görüntüler"""
        if self.mortality_table is None:
            print("❌ Önce mortalite tablosu oluşturulmalı!")
            return

        print(f"\n📋 Mortalite Tablosu (İlk {head} satır):")
        print("=" * 80)

        # Sayısal değerleri formatla
        display_df = self.mortality_table.head(head).copy()

        # Formatla
        for col in ['lx', 'dx', 'Lx', 'Tx']:
            if col in display_df.columns:
                display_df[col] = display_df[col].astype(float).round().astype(int)

        for col in ['qx', 'px', 'ex']:
            if col in display_df.columns:
                display_df[col] = display_df[col].astype(float).round(6)

        print(display_df.to_string(index=False))

    def save_to_excel(self, output_path="mortalite_tablosu.xlsx"):
        """Sonuçları Excel'e kaydet"""
        if self.mortality_table is None:
            print("❌ Kaydedilecek tablo yok!")
            return False

        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Ana tablo
                save_df = self.mortality_table.copy()

                # Sayısal kolonları düzenle
                for col in ['lx', 'dx', 'Lx', 'Tx']:
                    if col in save_df.columns:
                        save_df[col] = save_df[col].astype(float).round().astype(int)

                for col in ['qx', 'px', 'ex']:
                    if col in save_df.columns:
                        save_df[col] = save_df[col].astype(float).round(6)

                save_df.to_excel(writer, sheet_name='Mortalite_Tablosu', index=False)

                # Özet istatistikler
                summary_data = {
                    'İstatistik': ['Toplam Yaş Grubu', 'Ortalama Yaşam Beklentisi (0 yaş)',
                                   'En Yüksek Ölüm Oranı', 'En Düşük Ölüm Oranı (>0)', 'Medyan Yaşam Beklentisi'],
                    'Değer': [
                        len(self.mortality_table),
                        f"{self.mortality_table.iloc[0]['ex']:.2f} yıl" if 'ex' in self.mortality_table.columns else 'N/A',
                        f"{self.mortality_table['qx'].max():.6f}",
                        f"{self.mortality_table[self.mortality_table['qx'] > 0]['qx'].min():.6f}",
                        f"{self.mortality_table['ex'].median():.2f} yıl" if 'ex' in self.mortality_table.columns else 'N/A'
                    ]
                }
                summary = pd.DataFrame(summary_data)
                summary.to_excel(writer, sheet_name='Özet', index=False)

            print(f"✅ Mortalite tablosu kaydedildi: {output_path}")
            return True

        except Exception as e:
            print(f"❌ Kaydetme hatası: {str(e)}")
            return False

    def plot_mortality_curves(self):
        """Mortalite eğrilerini çizer"""
        if self.mortality_table is None:
            print("❌ Önce mortalite tablosu oluşturulmalı!")
            return

        try:
            # Türkçe karakter desteği için font ayarı
            plt.rcParams['font.family'] = ['DejaVu Sans', 'Arial', 'sans-serif']

            fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(15, 10))

            # 1. Yaşayan sayısı (lx)
            ax1.plot(self.mortality_table['yaş'], self.mortality_table['lx'], 'b-', linewidth=2)
            ax1.set_title('Yaşayan Sayısı (lx)', fontsize=12, fontweight='bold')
            ax1.set_xlabel('Yaş')
            ax1.set_ylabel('Yaşayan Sayısı')
            ax1.grid(True, alpha=0.3)
            ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{int(x):,}'))

            # 2. Ölüm oranı (qx)
            valid_qx = self.mortality_table[self.mortality_table['qx'] > 0]
            ax2.semilogy(valid_qx['yaş'], valid_qx['qx'], 'r-', linewidth=2)
            ax2.set_title('Ölüm Oranı (qx) - Log Ölçek', fontsize=12, fontweight='bold')
            ax2.set_xlabel('Yaş')
            ax2.set_ylabel('Ölüm Oranı (log)')
            ax2.grid(True, alpha=0.3)

            # 3. Ölen sayısı (dx)
            ax3.plot(self.mortality_table['yaş'], self.mortality_table['dx'], 'g-', linewidth=2)
            ax3.set_title('Ölen Sayısı (dx)', fontsize=12, fontweight='bold')
            ax3.set_xlabel('Yaş')
            ax3.set_ylabel('Ölen Sayısı')
            ax3.grid(True, alpha=0.3)
            ax3.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{int(x):,}'))

            # 4. Yaşam beklentisi (ex)
            if 'ex' in self.mortality_table.columns:
                ax4.plot(self.mortality_table['yaş'], self.mortality_table['ex'], 'm-', linewidth=2)
                ax4.set_title('Yaşam Beklentisi (ex)', fontsize=12, fontweight='bold')
                ax4.set_xlabel('Yaş')
                ax4.set_ylabel('Yaşam Beklentisi (yıl)')
                ax4.grid(True, alpha=0.3)
            else:
                ax4.text(0.5, 0.5, 'Yaşam Beklentisi\nHesaplanamadı',
                         horizontalalignment='center', verticalalignment='center',
                         transform=ax4.transAxes, fontsize=12)
                ax4.set_title('Yaşam Beklentisi (ex)', fontsize=12, fontweight='bold')

            plt.tight_layout()

            # Grafik dosyasını kaydet
            try:
                plt.savefig('mortalite_grafikleri.png', dpi=300, bbox_inches='tight')
                print("📊 Grafikler 'mortalite_grafikleri.png' olarak kaydedildi")
            except:
                pass

            plt.show()

        except Exception as e:
            print(f"❌ Grafik oluşturma hatası: {str(e)}")
            print("💡 İpucu: matplotlib kütüphanesinin yüklü olduğundan emin olun")


def create_sample_data_interactive():
    """Interaktif örnek veri oluşturucu"""
    print("\n📝 Hangi tür örnek veri oluşturmak istersiniz?")
    print("1. Sadece qx (ölüm oranları) - En basit")
    print("2. Yaş + qx + lx - Standart format")
    print("3. Tam veri (yaş, qx, lx, dx) - Eksiksiz")

    data_type = input("Seçiminiz (1-3): ").strip()

    ages = list(range(0, 101))

    try:
        if data_type == "1":
            # Sadece qx
            sample = pd.DataFrame({
                'yaş': ages,
                'qx': [0.007 + (x / 100) ** 2 * 0.1 for x in ages]
            })
            filename = "ornek_qx_verisi.xlsx"

        elif data_type == "2":
            # qx + lx
            qx_vals = [min(0.95, 0.007 + (x / 100) ** 2 * 0.1) for x in ages]
            lx_vals = [100000]

            for i in range(1, len(ages)):
                new_lx = max(0, int(lx_vals[i - 1] * (1 - qx_vals[i - 1])))
                lx_vals.append(new_lx)

            sample = pd.DataFrame({
                'yaş': ages,
                'qx': qx_vals,
                'lx': lx_vals
            })
            filename = "ornek_qx_lx_verisi.xlsx"

        else:
            # Tam veri
            qx_vals = [min(0.95, 0.007 + (x / 100) ** 2 * 0.1) for x in ages]
            lx_vals = [100000]
            dx_vals = []

            for i in range(len(ages)):
                if i == 0:
                    dx_vals.append(int(lx_vals[0] * qx_vals[0]))
                else:
                    new_lx = max(0, lx_vals[i - 1] - dx_vals[i - 1])
                    lx_vals.append(new_lx)
                    if i < len(ages) - 1:
                        dx_vals.append(int(new_lx * qx_vals[i]))
                    else:
                        dx_vals.append(new_lx)  # Son yaş grubunda herkes ölür

            sample = pd.DataFrame({
                'yaş': ages,
                'qx': qx_vals,
                'lx': lx_vals,
                'dx': dx_vals
            })
            filename = "ornek_tam_veri.xlsx"

        sample.to_excel(filename, index=False)
        print(f"✅ Örnek veri oluşturuldu: {filename}")
        return filename

    except Exception as e:
        print(f"❌ Örnek veri oluşturma hatası: {str(e)}")
        return None


def main():
    """Ana program"""
    print("🏥 Mortalite Tablosu Oluşturucu")
    print("=" * 50)

    # Mortalite tablosu oluşturucu
    mt_generator = MortalityTableGenerator()

    # Kullanıcıdan dosya yolu al
    print("\n📁 Excel Dosyası Seçenekleri:")
    print("1. Kendi Excel dosyanızın yolunu girin")
    print("2. Örnek veri oluştur ve test et")

    choice = input("\nSeçiminiz (1/2): ").strip()

    if choice == "1":
        excel_file = input("Excel dosyasının tam yolunu girin: ").strip().replace('"', '')

        # Dosya var mı kontrol et
        if not Path(excel_file).exists():
            print(f"❌ Dosya bulunamadı: {excel_file}")
            return

    else:
        # Örnek veri oluştur
        print("\n📝 Örnek veri oluşturuluyor...")
        excel_file = create_sample_data_interactive()

        if excel_file is None:
            print("❌ Örnek veri oluşturulamadı!")
            return

    print(f"\n📖 Excel dosyası okunuyor: {excel_file}")

    # Veriyi oku
    if mt_generator.read_excel_data(excel_file) is not None:

        # Radix değeri al
        radix_input = input(f"\nBaşlangıç yaşayan sayısı (varsayılan: 100000): ").strip()
        radix = int(radix_input) if radix_input.isdigit() else 100000

        print(f"\n⚙  Mortalite tablosu oluşturuluyor (radix={radix:,})...")

        # Mortalite tablosunu oluştur
        mortality_table = mt_generator.create_mortality_table(radix=radix)

        if mortality_table is not None:
            print(f"\n✅ Başarıyla {len(mortality_table)} yaş grubu için tablo oluşturuldu!")

            # Tabloyu göster
            print(f"\n📋 Tablo önizlemesi:")
            mt_generator.display_table(head=10)

            # Özet istatistikler
            print(f"\n📊 Özet İstatistikler:")
            print(f"• Yaş aralığı: {int(mortality_table['yaş'].min())}-{int(mortality_table['yaş'].max())}")
            print(f"• Ortalama ölüm oranı: {mortality_table['qx'].mean():.6f}")
            max_qx_age = int(mortality_table.loc[mortality_table['qx'].idxmax(), 'yaş'])
            print(f"• En yüksek ölüm oranı: {mortality_table['qx'].max():.6f} ({max_qx_age} yaş)")

            if 'ex' in mortality_table.columns and not pd.isna(mortality_table.iloc[0]['ex']):
                print(f"• 0 yaş yaşam beklentisi: {mortality_table.iloc[0]['ex']:.2f} yıl")
                age_65_data = mortality_table[mortality_table['yaş'] == 65]
                if len(age_65_data) > 0:
                    ex_65 = age_65_data['ex'].iloc[0]
                    if not pd.isna(ex_65):
                        print(f"• 65 yaş yaşam beklentisi: {ex_65:.2f} yıl")

            # Kaydetme seçenekleri
            print(f"\n💾 Sonuçları kaydetmek ister misiniz?")
            save_choice = input("Excel dosyası olarak kaydet? (e/h): ").strip().lower()

            if save_choice in ['e', 'evet', 'y', 'yes']:
                output_file = input("Çıktı dosya adı (varsayılan: mortalite_tablosu.xlsx): ").strip()
                if not output_file:
                    output_file = "mortalite_tablosu.xlsx"

                mt_generator.save_to_excel(output_file)

            # Grafik seçenekleri
            print(f"\n📈 Grafikleri görüntülemek ister misiniz?")
            plot_choice = input("Mortalite eğrilerini çiz? (e/h): ").strip().lower()

            if plot_choice in ['e', 'evet', 'y', 'yes']:
                print("📊 Grafikler oluşturuluyor...")
                mt_generator.plot_mortality_curves()

            print(f"\n🎉 İşlem tamamlandı!")

        else:
            print("❌ Mortalite tablosu oluşturulamadı!")

    else:
        print("❌ Excel dosyası okunamadı. Lütfen dosya yolunu ve formatını kontrol edin.")
        print("\n📋 Beklenen Excel formatı:")
        print("• 'yaş' veya 'age' sütunu: 0, 1, 2, ... yaş değerleri")
        print("• 'qx' sütunu: Ölüm oranları (0-1 arası)")
        print("• 'lx' sütunu: Yaşayan sayıları (opsiyonel)")
        print("• 'dx' sütunu: Ölen sayıları (opsiyonel)")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n👋 Program kullanıcı tarafından sonlandırıldı.")
    except Exception as e:
        print(f"\n❌ Beklenmeyen hata: {str(e)}")
        print("\n🔧 Çözüm önerileri:")
        print("1. Gerekli kütüphaneler yüklü mü kontrol edin:")
        print("   pip install pandas numpy matplotlib openpyxl")
        print("2. Excel dosyasının formatını kontrol edin")
        print("3. Dosya yolunda Türkçe karakter varsa İngilizce yol deneyin")