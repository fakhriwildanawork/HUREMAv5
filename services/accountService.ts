
import { supabase } from '../lib/supabase';
import { Account, AccountInput, CareerLog, CareerLogInput, HealthLog, HealthLogInput } from '../types';

/**
 * Fungsi pembantu untuk membersihkan data sebelum dikirim ke Supabase.
 * Mengubah string kosong ('') menjadi null agar tidak error saat masuk ke kolom UUID atau DATE.
 */
const sanitizePayload = (payload: any) => {
  const sanitized: any = {};
  Object.keys(payload).forEach(key => {
    if (payload[key] === undefined) return;
    sanitized[key] = payload[key] === '' ? null : payload[key];
  });
  return sanitized;
};

export const accountService = {
  async getAll() {
    const { data, error } = await supabase
      .from('accounts')
      .select(`
        *,
        location:locations(name)
      `)
      .not('access_code', 'ilike', '%SPADMIN%')
      .order('created_at', { ascending: false });
    
    if (error) {
      console.error("SUPABASE_GET_ALL_ERROR:", error.message);
      throw error;
    }
    return data;
  },

  async getById(id: string) {
    const { data, error } = await supabase
      .from('accounts')
      .select(`
        *,
        location:locations(*),
        schedule:schedules!schedule_id(*, rules:schedule_rules(*))
      `)
      .eq('id', id)
      .maybeSingle(); // Menggunakan maybeSingle() agar tidak error jika record tidak ditemukan
    
    if (error) {
      console.error("SUPABASE_GET_BY_ID_ERROR:", error.message);
      throw error;
    }
    return data;
  },

  async getDistinctAttributes() {
    // Mengambil data unik Jabatan & Golongan dari kedua tabel untuk memastikan dropdown lengkap
    const [accRes, logRes] = await Promise.all([
      supabase.from('accounts').select('position, grade'),
      supabase.from('account_career_logs').select('position, grade')
    ]);

    const allPositions = [
      ...(accRes.data?.map(p => p.position) || []),
      ...(logRes.data?.map(p => p.position) || [])
    ];
    
    const allGrades = [
      ...(accRes.data?.map(g => g.grade) || []),
      ...(logRes.data?.map(g => g.grade) || [])
    ];

    const uniquePositions = Array.from(new Set(allPositions.filter(Boolean))).sort();
    const uniqueGrades = Array.from(new Set(allGrades.filter(Boolean))).sort();
    
    return { positions: uniquePositions, grades: uniqueGrades };
  },

  async getCareerLogs(accountId: string) {
    const { data, error } = await supabase
      .from('account_career_logs')
      .select('*')
      .eq('account_id', accountId)
      .order('change_date', { ascending: false });
    
    if (error) throw error;
    return data as CareerLog[];
  },

  async getHealthLogs(accountId: string) {
    const { data, error } = await supabase
      .from('account_health_logs')
      .select('*')
      .eq('account_id', accountId)
      .order('change_date', { ascending: false });
    
    if (error) throw error;
    return data as HealthLog[];
  },

  async create(account: AccountInput & { file_sk_id?: string, file_mcu_id?: string, contract_initial?: any }) {
    const { file_sk_id, file_mcu_id, contract_initial, ...rest } = account;
    
    const sanitizedAccount = sanitizePayload(rest);
    
    // 1. Insert ke tabel accounts
    const { data, error } = await supabase
      .from('accounts')
      .insert([sanitizedAccount])
      .select();
    
    if (error) {
      console.error("SUPABASE_CREATE_ERROR:", error.message);
      throw error;
    }
    
    const newAccount = data[0] as Account;

    // 2. Otomatis buat log karier awal
    const { data: locData } = await supabase
      .from('locations')
      .select('name')
      .eq('id', newAccount.location_id)
      .single();

    await supabase.from('account_career_logs').insert([{
      account_id: newAccount.id,
      position: newAccount.position,
      grade: newAccount.grade,
      location_id: newAccount.location_id,
      location_name: locData?.name || '-',
      schedule_id: newAccount.schedule_id,
      file_sk_id: file_sk_id || null,
      notes: 'Initial Career Record'
    }]);

    // 3. Otomatis buat log kesehatan awal
    await supabase.from('account_health_logs').insert([{
      account_id: newAccount.id,
      file_mcu_id: file_mcu_id || null,
      notes: 'Initial Health Record'
    }]);

    // 4. Otomatis buat kontrak awal jika disediakan
    if (contract_initial && contract_initial.contract_number) {
      await supabase.from('account_contracts').insert([{
        account_id: newAccount.id,
        contract_number: contract_initial.contract_number,
        contract_type: contract_initial.contract_type || (account.employee_type === 'Kontrak' ? 'PKWT' : account.employee_type),
        start_date: contract_initial.start_date || account.start_date,
        end_date: contract_initial.end_date || account.end_date,
        file_id: contract_initial.file_id || null,
        notes: 'Initial Contract Record'
      }]);
    }

    return newAccount;
  },

  async update(id: string, account: Partial<AccountInput>) {
    // Pastikan field tambahan untuk log awal tidak ikut dikirim ke tabel accounts
    const { file_sk_id, file_mcu_id, contract_initial, ...rest } = account as any;
    
    const sanitizedAccount = sanitizePayload(rest);
    
    const { data, error } = await supabase
      .from('accounts')
      .update(sanitizedAccount)
      .eq('id', id)
      .select(`
        *,
        location:locations(name)
      `);
    
    if (error) {
      console.error("SUPABASE_UPDATE_ERROR:", error.message);
      throw error;
    }
    return data[0] as Account;
  },

  async delete(id: string) {
    // 1. Ambil data akun untuk mendapatkan ID file Drive
    const { data: account } = await supabase
      .from('accounts')
      .select('photo_google_id, ktp_google_id, diploma_google_id')
      .eq('id', id)
      .single();

    // 2. Ambil semua log untuk mendapatkan ID file Drive mereka
    const [careerLogs, healthLogs, contracts, warnings, terminations] = await Promise.all([
      supabase.from('account_career_logs').select('file_sk_id').eq('account_id', id),
      supabase.from('account_health_logs').select('file_mcu_id').eq('account_id', id),
      supabase.from('account_contracts').select('file_id').eq('account_id', id),
      supabase.from('account_warning_logs').select('file_id').eq('account_id', id),
      supabase.from('account_termination_logs').select('file_id').eq('account_id', id)
    ]);

    // 3. Kumpulkan semua file ID unik
    const fileIds = new Set<string>();
    if (account?.photo_google_id) fileIds.add(account.photo_google_id);
    if (account?.ktp_google_id) fileIds.add(account.ktp_google_id);
    if (account?.diploma_google_id) fileIds.add(account.diploma_google_id);
    
    careerLogs.data?.forEach(log => log.file_sk_id && fileIds.add(log.file_sk_id));
    healthLogs.data?.forEach(log => log.file_mcu_id && fileIds.add(log.file_mcu_id));
    contracts.data?.forEach(log => log.file_id && fileIds.add(log.file_id));
    warnings.data?.forEach(log => log.file_id && fileIds.add(log.file_id));
    terminations.data?.forEach(log => log.file_id && fileIds.add(log.file_id));

    // 4. Hapus file dari Google Drive secara paralel (dengan sedikit delay/batching jika banyak)
    const { googleDriveService } = await import('./googleDriveService');
    for (const fileId of fileIds) {
      try {
        await googleDriveService.deleteFile(fileId);
      } catch (e) {
        console.error(`Gagal menghapus file ${fileId} saat hapus akun:`, e);
      }
    }

    // 5. Hapus dari database (Supabase akan handle cascade jika diatur, tapi kita pastikan)
    const { error } = await supabase
      .from('accounts')
      .delete()
      .eq('id', id);
    
    if (error) {
      console.error("SUPABASE_DELETE_ERROR:", error.message);
      throw error;
    }
    return true;
  },

  async bulkCreate(accounts: (AccountInput & { location_name?: string, schedule_name?: string })[]) {
    const results = {
      success: 0,
      failed: 0,
      errors: [] as string[]
    };

    // 1. Ambil semua lokasi dan jadwal sekaligus untuk pemetaan (Caching)
    const [{ data: locations }, { data: schedules }] = await Promise.all([
      supabase.from('locations').select('id, name'),
      supabase.from('schedules').select('id, name, type')
    ]);

    const locationMap = new Map(locations?.map(l => [l.name.toLowerCase(), l.id]));
    const locationNameMap = new Map(locations?.map(l => [l.id, l.name]));
    const scheduleMap = new Map(schedules?.map(s => [s.name.toLowerCase(), s]));

    // 2. Proses dalam batch paralel untuk kecepatan (Concurrency: 5)
    const batchSize = 5;
    for (let i = 0; i < accounts.length; i += batchSize) {
      const batch = accounts.slice(i, i + batchSize);
      
      await Promise.all(batch.map(async (acc) => {
        try {
          const { location_name, schedule_name, ...rest } = acc;
          
          // Map location name to ID
          if (!rest.location_id && location_name) {
            const locId = locationMap.get(location_name.toLowerCase());
            if (!locId) {
              throw new Error(`Lokasi "${location_name}" tidak ditemukan.`);
            }
            rest.location_id = locId;
          } else if (!rest.location_id) {
            throw new Error(`Kolom Lokasi Penempatan wajib diisi.`);
          }

          // Map schedule name to ID and Type
          if (schedule_name) {
            const lowerName = schedule_name.toLowerCase();
            if (lowerName === 'fleksibel') {
              rest.schedule_type = 'FLEKSIBEL';
              rest.schedule_id = null;
            } else if (lowerName === 'shift dinamis') {
              rest.schedule_type = 'DINAMIS';
              rest.schedule_id = null;
            } else {
              const sch = scheduleMap.get(lowerName);
              if (!sch) {
                throw new Error(`Jadwal "${schedule_name}" tidak ditemukan.`);
              }
              rest.schedule_id = sch.id;
              rest.schedule_type = sch.type === 1 ? 'Office Hour' : 'Shift';
            }
          }

          // Gunakan versi optimasi yang tidak melakukan query lokasi ulang
          const locName = locationNameMap.get(rest.location_id) || '-';
          await this.createWithPredefinedLocation(rest, locName);
          results.success++;
        } catch (err: any) {
          results.failed++;
          results.errors.push(`${acc.full_name || 'Tanpa Nama'}: ${err.message}`);
        }
      }));
    }

    return results;
  },

  /**
   * Versi optimasi dari create() khusus untuk bulk import.
   * Menghindari query SELECT tambahan untuk nama lokasi karena sudah ada di Map.
   */
  async createWithPredefinedLocation(account: AccountInput & { file_sk_id?: string, file_mcu_id?: string, contract_initial?: any }, locationName: string) {
    const { file_sk_id, file_mcu_id, contract_initial, ...rest } = account;
    const sanitizedAccount = sanitizePayload(rest);
    
    // 1. Insert ke tabel accounts
    const { data, error } = await supabase
      .from('accounts')
      .insert([sanitizedAccount])
      .select();
    
    if (error) throw error;
    const newAccount = data[0] as Account;

    // 2. Buat log & kontrak secara paralel
    const promises = [
      supabase.from('account_career_logs').insert([{
        account_id: newAccount.id,
        position: newAccount.position,
        grade: newAccount.grade,
        location_id: newAccount.location_id,
        location_name: locationName,
        schedule_id: newAccount.schedule_id,
        file_sk_id: file_sk_id || null,
        notes: 'Initial Career Record'
      }]),
      supabase.from('account_health_logs').insert([{
        account_id: newAccount.id,
        file_mcu_id: file_mcu_id || null,
        notes: 'Initial Health Record'
      }])
    ];

    if (contract_initial && contract_initial.contract_number) {
      promises.push(supabase.from('account_contracts').insert([{
        account_id: newAccount.id,
        contract_number: contract_initial.contract_number,
        contract_type: contract_initial.contract_type || (account.employee_type === 'Kontrak' ? 'PKWT' : account.employee_type),
        start_date: contract_initial.start_date || account.start_date,
        end_date: contract_initial.end_date || account.end_date,
        file_id: contract_initial.file_id || null,
        notes: 'Initial Contract Record'
      }]));
    }

    await Promise.all(promises);
    return newAccount;
  },

  async updateImageByNikOrName(identifier: string, type: 'photo' | 'ktp' | 'ijazah' | 'sk' | 'mcu' | 'kontrak', fileId: string) {
    // Normalize identifier for name matching
    const normalizedId = identifier.toLowerCase().replace(/[^a-z0-9]/g, '');
    
    // 1. Try matching by internal_nik first
    const { data: byNik, error: nikError } = await supabase
      .from('accounts')
      .select('id, full_name')
      .eq('internal_nik', identifier)
      .maybeSingle();

    if (nikError) throw nikError;

    let targetId = byNik?.id;
    let targetName = byNik?.full_name;

    if (!targetId) {
      // 2. Try matching by full_name (normalized)
      const { data: allAccounts, error: allErr } = await supabase.from('accounts').select('id, full_name');
      if (allErr) throw allErr;

      const matches = allAccounts?.filter(acc => 
        acc.full_name.toLowerCase().replace(/[^a-z0-9]/g, '') === normalizedId
      );

      if (!matches || matches.length === 0) {
        throw new Error(`Akun dengan Nama/NIK "${identifier}" tidak ditemukan.`);
      }

      if (matches.length > 1) {
        throw new Error(`Nama "${identifier}" ditemukan lebih dari satu akun (${matches.length}). Gunakan NIK agar spesifik.`);
      }

      targetId = matches[0].id;
      targetName = matches[0].full_name;
    }

    // 3. Update field sesuai tipe
    if (type === 'photo' || type === 'ktp' || type === 'ijazah') {
      const updateData: any = {};
      if (type === 'photo') updateData.photo_google_id = fileId;
      else if (type === 'ktp') updateData.ktp_google_id = fileId;
      else if (type === 'ijazah') updateData.diploma_google_id = fileId;

      const { error } = await supabase.from('accounts').update(updateData).eq('id', targetId);
      if (error) throw error;
    } else if (type === 'sk') {
      const { data: latestLog } = await supabase
        .from('account_career_logs')
        .select('id')
        .eq('account_id', targetId)
        .order('change_date', { ascending: false })
        .limit(1);
      if (latestLog && latestLog.length > 0) {
        await supabase.from('account_career_logs').update({ file_sk_id: fileId }).eq('id', latestLog[0].id);
      }
    } else if (type === 'mcu') {
      const { data: latestLog } = await supabase
        .from('account_health_logs')
        .select('id')
        .eq('account_id', targetId)
        .order('change_date', { ascending: false })
        .limit(1);
      if (latestLog && latestLog.length > 0) {
        await supabase.from('account_health_logs').update({ file_mcu_id: fileId }).eq('id', latestLog[0].id);
      }
    } else if (type === 'kontrak') {
      const { data: latestContract } = await supabase
        .from('account_contracts')
        .select('id')
        .eq('account_id', targetId)
        .order('start_date', { ascending: false })
        .limit(1);
      if (latestContract && latestContract.length > 0) {
        await supabase.from('account_contracts').update({ file_id: fileId }).eq('id', latestContract[0].id);
      }
    }

    return { success: true, name: targetName };
  },

  // Manual Log Management
  async createCareerLog(logInput: CareerLogInput) {
    // Filtrasi: Pastikan hanya kolom yang ada di tabel account_career_logs yang dikirim
    const { account_id, position, grade, location_name, file_sk_id, notes, location_id, schedule_id, schedule_type, change_date } = logInput;
    const payload = sanitizePayload({ account_id, position, grade, location_name, file_sk_id, notes, change_date, location_id, schedule_id, schedule_type });
    
    const { data, error } = await supabase
      .from('account_career_logs')
      .insert([payload])
      .select();

    if (error) {
      console.error("CAREER_LOG_CREATE_ERROR:", error.message);
      throw error;
    }

    // Sinkronisasi ke profil utama jika data karier berubah
    await this.update(account_id, {
      position,
      grade,
      location_id: location_id || null,
      schedule_id: schedule_id || null,
      schedule_type: schedule_type || null
    });

    return data[0] as CareerLog;
  },

  async updateCareerLog(id: string, logInput: Partial<CareerLogInput>) {
    const { account_id, position, grade, location_name, file_sk_id, notes, location_id, schedule_id, schedule_type, change_date } = logInput;
    const payload = sanitizePayload({ account_id, position, grade, location_name, file_sk_id, notes, change_date, location_id, schedule_id, schedule_type });

    const { data, error } = await supabase
      .from('account_career_logs')
      .update(payload)
      .eq('id', id)
      .select();

    if (error) {
      console.error("CAREER_LOG_UPDATE_ERROR:", error.message);
      throw error;
    }

    if (account_id) {
      await this.update(account_id, {
        position,
        grade,
        location_id: location_id || null,
        schedule_id: schedule_id || null,
        schedule_type: schedule_type || null
      });
    }

    return data[0] as CareerLog;
  },

  async deleteCareerLog(id: string) {
    // 1. Ambil ID file
    const { data } = await supabase.from('account_career_logs').select('file_sk_id').eq('id', id).single();
    
    // 2. Hapus file dari Drive
    if (data?.file_sk_id) {
      const { googleDriveService } = await import('./googleDriveService');
      await googleDriveService.deleteFile(data.file_sk_id);
    }

    // 3. Hapus dari DB
    const { error } = await supabase
      .from('account_career_logs')
      .delete()
      .eq('id', id);
    if (error) throw error;
    return true;
  },

  async createHealthLog(logInput: HealthLogInput) {
    // Filtrasi: Hapus field career (location_id, location_name) yang sering terbawa dari state form
    const { account_id, mcu_status, health_risk, file_mcu_id, notes, change_date } = logInput;
    const payload = sanitizePayload({ account_id, mcu_status, health_risk, file_mcu_id, notes, change_date });

    const { data, error } = await supabase
      .from('account_health_logs')
      .insert([payload])
      .select();

    if (error) {
      console.error("HEALTH_LOG_CREATE_ERROR:", error.message);
      throw error;
    }

    // Sinkronisasi ke profil utama
    await this.update(account_id, {
      mcu_status,
      health_risk
    });

    return data[0] as HealthLog;
  },

  async updateHealthLog(id: string, logInput: Partial<HealthLogInput>) {
    const { account_id, mcu_status, health_risk, file_mcu_id, notes, change_date } = logInput;
    const payload = sanitizePayload({ account_id, mcu_status, health_risk, file_mcu_id, notes, change_date });

    const { data, error } = await supabase
      .from('account_health_logs')
      .update(payload)
      .eq('id', id)
      .select();

    if (error) {
      console.error("HEALTH_LOG_UPDATE_ERROR:", error.message);
      throw error;
    }

    if (account_id) {
      await this.update(account_id, {
        mcu_status,
        health_risk
      });
    }

    return data[0] as HealthLog;
  },

  async deleteHealthLog(id: string) {
    // 1. Ambil ID file
    const { data } = await supabase.from('account_health_logs').select('file_mcu_id').eq('id', id).single();
    
    // 2. Hapus file dari Drive
    if (data?.file_mcu_id) {
      const { googleDriveService } = await import('./googleDriveService');
      await googleDriveService.deleteFile(data.file_mcu_id);
    }

    // 3. Hapus dari DB
    const { error } = await supabase
      .from('account_health_logs')
      .delete()
      .eq('id', id);
    if (error) throw error;
    return true;
  }
};
