import { supabase } from '../lib/supabase';
import { AppSetting } from '../types';

export const settingsService = {
  /**
   * Mendapatkan nilai pengaturan berdasarkan key
   */
  async getSetting(key: string, defaultValue: any = null): Promise<any> {
    try {
      const { data, error } = await supabase
        .from('app_settings')
        .select('value')
        .eq('key', key)
        .maybeSingle();
      
      if (error) throw error;
      return data ? data.value : defaultValue;
    } catch (error) {
      console.error(`Error getting setting ${key}:`, error);
      return defaultValue;
    }
  },

  /**
   * Mendapatkan semua pengaturan
   */
  async getAll(): Promise<AppSetting[]> {
    try {
      const { data, error } = await supabase
        .from('app_settings')
        .select('*');
      
      if (error) {
        // Jika tabel tidak ditemukan (PGRST205), kembalikan array kosong tanpa melempar error
        if (error.code === 'PGRST205') return [];
        throw error;
      }
      return data as AppSetting[];
    } catch (error) {
      console.error("Error fetching all settings:", error);
      return [];
    }
  },

  /**
   * Memperbarui atau membuat pengaturan baru
   */
  async updateSetting(key: string, value: any, description?: string) {
    try {
      const { error } = await supabase
        .from('app_settings')
        .upsert({ 
          key, 
          value, 
          description, 
          updated_at: new Date().toISOString() 
        });
      
      if (error) throw error;
      return true;
    } catch (error: any) {
      console.error("Error updating setting:", error);
      throw error;
    }
  }
};