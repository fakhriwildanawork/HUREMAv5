import { supabase } from '../lib/supabase';
import { Meeting, MeetingInput, MeetingNote } from '../types';

export const meetingService = {
  async getAll(): Promise<Meeting[]> {
    const { data, error } = await supabase
      .from('meetings')
      .select('*, creator:accounts!created_by(full_name)')
      .order('scheduled_at', { ascending: false });
    
    if (error) throw error;

    // Fetch details for participants and notulens
    const meetingsWithDetails = await Promise.all((data || []).map(async (meeting) => {
      const [participantsRes, notulensRes, notesRes] = await Promise.all([
        meeting.participant_ids?.length > 0 
          ? supabase.from('accounts').select('id, full_name, internal_nik').in('id', meeting.participant_ids)
          : Promise.resolve({ data: [] }),
        meeting.notulen_ids?.length > 0
          ? supabase.from('accounts').select('id, full_name, internal_nik').in('id', meeting.notulen_ids)
          : Promise.resolve({ data: [] }),
        supabase.from('meeting_notes').select('*').eq('meeting_id', meeting.id).order('created_at', { ascending: true })
      ]);

      return { 
        ...meeting, 
        participants: participantsRes.data || [],
        notulens: notulensRes.data || [],
        notes: notesRes.data || []
      };
    }));

    return meetingsWithDetails;
  },

  async getById(id: string): Promise<Meeting & { notes: MeetingNote[] }> {
    const { data, error } = await supabase
      .from('meetings')
      .select('*, creator:accounts!created_by(full_name)')
      .eq('id', id)
      .single();
    
    if (error) throw error;

    const [participantsRes, notulensRes, notesRes] = await Promise.all([
      data.participant_ids?.length > 0 
        ? supabase.from('accounts').select('id, full_name, internal_nik').in('id', data.participant_ids)
        : Promise.resolve({ data: [] }),
      data.notulen_ids?.length > 0
        ? supabase.from('accounts').select('id, full_name, internal_nik').in('id', data.notulen_ids)
        : Promise.resolve({ data: [] }),
      supabase.from('meeting_notes').select('*').eq('meeting_id', id).order('created_at', { ascending: true })
    ]);

    return { 
      ...data, 
      participants: participantsRes.data || [],
      notulens: notulensRes.data || [],
      notes: notesRes.data || []
    };
  },

  async create(input: MeetingInput): Promise<void> {
    const { error } = await supabase
      .from('meetings')
      .insert([input]);
    
    if (error) throw error;
  },

  async update(id: string, input: Partial<MeetingInput>): Promise<void> {
    const { error } = await supabase
      .from('meetings')
      .update(input)
      .eq('id', id);
    
    if (error) throw error;
  },

  async delete(id: string): Promise<void> {
    const { error } = await supabase
      .from('meetings')
      .delete()
      .eq('id', id);
    
    if (error) throw error;
  },

  async startMeeting(id: string): Promise<void> {
    const { error } = await supabase
      .from('meetings')
      .update({ 
        status: 'In Progress',
        started_at: new Date().toISOString()
      })
      .eq('id', id);
    
    if (error) throw error;
  },

  async endMeeting(id: string): Promise<void> {
    const { error } = await supabase
      .from('meetings')
      .update({ 
        status: 'Completed',
        ended_at: new Date().toISOString()
      })
      .eq('id', id);
    
    if (error) throw error;
  },

  // Meeting Notes Methods
  async getNotes(meetingId: string): Promise<MeetingNote[]> {
    const { data, error } = await supabase
      .from('meeting_notes')
      .select('*')
      .eq('meeting_id', meetingId)
      .order('created_at', { ascending: true });
    
    if (error) throw error;
    return data || [];
  },

  async addNote(meetingId: string, content: string, attachments: string[], links: string[]): Promise<void> {
    const { error } = await supabase
      .from('meeting_notes')
      .insert([{
        meeting_id: meetingId,
        content,
        attachments,
        links
      }]);
    
    if (error) throw error;
  },

  async updateNote(id: string, content: string, attachments: string[], links: string[]): Promise<void> {
    const { error } = await supabase
      .from('meeting_notes')
      .update({
        content,
        attachments,
        links,
        updated_at: new Date().toISOString()
      })
      .eq('id', id);
    
    if (error) throw error;
  },

  async deleteNote(id: string): Promise<void> {
    const { error } = await supabase
      .from('meeting_notes')
      .delete()
      .eq('id', id);
    
    if (error) throw error;
  }
};
