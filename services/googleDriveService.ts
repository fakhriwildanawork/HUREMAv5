
class GoogleDriveService {
  /**
   * uploadQueue digunakan untuk memastikan file diunggah satu per satu (Sequential).
   * Ini krusial untuk mencegah error 401 (Unauthorized) dari Google OAuth saat
   * beberapa request mencoba melakukan refresh token di saat yang bersamaan.
   */
  private uploadQueue: Promise<any> = Promise.resolve();

  async uploadFile(file: File, folderId?: string): Promise<string> {
    // Memasukkan proses upload ke dalam antrean (Promise Chain)
    const currentUpload = this.uploadQueue.then(async () => {
      try {
        const formData = new FormData();
        formData.append('file', file);
        if (folderId) {
          formData.append('folderId', folderId);
        }

        const response = await fetch('/api/upload', {
          method: 'POST',
          body: formData,
        });

        if (!response.ok) {
          const errorData = await response.json().catch(() => ({}));
          throw new Error(errorData.error || 'Gagal mengunggah file. Periksa koneksi atau kredensial Google Drive.');
        }

        const result = await response.json();
        if (!result.id) {
          throw new Error('ID File tidak ditemukan dalam respon API Google Drive');
        }

        return `${result.id}|${file.name}`;
      } catch (error) {
        console.error('GoogleDriveService Upload Error:', error);
        throw error;
      }
    });

    // Perbarui antrean agar request berikutnya menunggu proses ini selesai
    this.uploadQueue = currentUpload.catch(() => {
      // Jika upload ini gagal, pastikan antrean tetap berlanjut untuk file berikutnya
      return null;
    });

    return currentUpload;
  }

  /**
   * Mendapatkan URL file dari ID File Google Drive.
   * Jika formatnya 'id|filename', akan dicek ekstensinya.
   * Jika gambar, gunakan lh3.googleusercontent. Jika bukan, gunakan Google Drive viewer.
   */
  getFileUrl(fileId: string, fullSize: boolean = false): string {
    if (!fileId) return '';
    
    let id = fileId;
    let name = '';
    if (fileId.includes('|')) {
      [id, name] = fileId.split('|');
    }

    // Jika tidak ada nama (data lama), default ke image untuk keamanan tampilan profil/presensi
    // Namun untuk kontrak, kita akan menangani di level UI jika perlu default ke document
    const isImage = !name || /\.(jpg|jpeg|png|webp|gif|svg|bmp)$/i.test(name);
    
    if (isImage) {
      return `https://lh3.googleusercontent.com/d/${id}=s${fullSize ? '0' : '1600'}`;
    }
    // Untuk non-image, buka via Google Drive viewer
    return `https://drive.google.com/file/d/${id}/view`;
  }

  /**
   * Menghapus file secara permanen dari Google Drive.
   */
  async deleteFile(fileId: string): Promise<boolean> {
    if (!fileId) return true;
    const id = fileId.includes('|') ? fileId.split('|')[0] : fileId;
    
    try {
      const response = await fetch(`/api/delete-file?fileId=${id}`, {
        method: 'DELETE',
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        console.error('GoogleDriveService Delete Error:', errorData);
        // Tetap return true jika file tidak ditemukan (404) agar proses hapus di DB tetap lanjut
        if (response.status === 404) return true;
        throw new Error(errorData.error || 'Gagal menghapus file dari Google Drive.');
      }

      return true;
    } catch (error) {
      console.error('GoogleDriveService Delete Error:', error);
      throw error;
    }
  }

  /**
   * Mendapatkan URL viewer Google Drive langsung (untuk dibuka di tab baru).
   */
  getViewerUrl(fileId: string): string {
    if (!fileId) return '';
    const id = fileId.includes('|') ? fileId.split('|')[0] : fileId;
    return `https://drive.google.com/file/d/${id}/view`;
  }
}

export const googleDriveService = new GoogleDriveService();
