import { supabase } from '../lib/supabase';
import { Announcement, AnnouncementRead } from '../types';

export const announcementService = {
  async getAnnouncements(userId: string, department?: string) {
    const now = new Date().toISOString();
    
    // Fetch announcements that are active and targeted to the user
    // We fetch all and filter in JS for complex targeting logic or use a smart query
    const { data, error } = await supabase
      .from('announcements')
      .select(`
        *,
        creator:accounts!created_by(full_name)
      `)
      .lte('publish_start', now)
      .gte('publish_end', now)
      .order('created_at', { ascending: false });

    if (error) throw error;

    // Filter based on targeting
    const filtered = (data as any[]).filter(ann => {
      if (ann.target_type === 'All') return true;
      if (ann.target_type === 'Department' && department) {
        return ann.target_ids.includes(department);
      }
      if (ann.target_type === 'Individual') {
        return ann.target_ids.includes(userId);
      }
      return false;
    });

    // Check read status for each
    const { data: reads, error: readError } = await supabase
      .from('announcement_reads')
      .select('announcement_id')
      .eq('user_id', userId);

    if (readError) throw readError;

    const readIds = new Set(reads.map(r => r.announcement_id));

    return filtered.map(ann => ({
      ...ann,
      is_read: readIds.has(ann.id)
    })) as Announcement[];
  },

  async getAllAnnouncementsAdmin() {
    const { data, error } = await supabase
      .from('announcements')
      .select(`
        *,
        creator:accounts!created_by(full_name)
      `)
      .order('created_at', { ascending: false });

    if (error) throw error;

    // Get read counts
    const { data: readCounts, error: countError } = await supabase
      .from('announcement_reads')
      .select('announcement_id');

    if (countError) throw countError;

    const counts: Record<string, number> = {};
    readCounts.forEach(r => {
      counts[r.announcement_id] = (counts[r.announcement_id] || 0) + 1;
    });

    return data.map(ann => ({
      ...ann,
      read_count: counts[ann.id] || 0
    })) as Announcement[];
  },

  async createAnnouncement(announcement: Omit<Announcement, 'id' | 'created_at' | 'updated_at' | 'creator'>) {
    const { data, error } = await supabase
      .from('announcements')
      .insert([announcement])
      .select()
      .single();

    if (error) throw error;
    return data;
  },

  async updateAnnouncement(id: string, announcement: Partial<Announcement>) {
    const { data, error } = await supabase
      .from('announcements')
      .update({ ...announcement, updated_at: new Date().toISOString() })
      .eq('id', id)
      .select()
      .single();

    if (error) throw error;
    return data;
  },

  async deleteAnnouncement(id: string) {
    const { error } = await supabase
      .from('announcements')
      .delete()
      .eq('id', id);

    if (error) throw error;
  },

  async markAsRead(announcementId: string, userId: string) {
    const { error } = await supabase
      .from('announcement_reads')
      .upsert({ announcement_id: announcementId, user_id: userId, read_at: new Date().toISOString() }, { onConflict: 'announcement_id,user_id' });

    if (error) throw error;
  },

  async getReadReceipts(announcementId: string) {
    const { data, error } = await supabase
      .from('announcement_reads')
      .select(`
        *,
        user:accounts!user_id(full_name, internal_nik)
      `)
      .eq('announcement_id', announcementId);

    if (error) throw error;
    return data;
  }
};
