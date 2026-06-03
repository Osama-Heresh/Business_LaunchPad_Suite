import React, { useState, useEffect } from 'react';
import { Plus, Trash2, User, Users, Mail, X } from 'lucide-react';
import { supabase } from '../services/supabaseClient';

interface TeamMember {
  id: string;
  user_id: string;
  email: string;
  name: string;
  role: 'owner' | 'member';
  joined_at: string;
}

interface Team {
  id: string;
  name: string;
  created_at: string;
}

const TeamManagement: React.FC = () => {
  const [teams, setTeams] = useState<Team[]>([]);
  const [members, setMembers] = useState<TeamMember[]>([]);
  const [selectedTeam, setSelectedTeam] = useState<string | null>(null);
  const [showAddMember, setShowAddMember] = useState(false);
  const [showCreateTeam, setShowCreateTeam] = useState(false);
  const [newTeamName, setNewTeamName] = useState('');
  const [newMemberEmail, setNewMemberEmail] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState('');

  const currentUserId = localStorage.getItem('user_id');

  useEffect(() => {
    if (currentUserId) {
      fetchTeams();
    }
  }, [currentUserId]);

  const fetchTeams = async () => {
    try {
      setLoading(true);
      const { data, error } = await supabase
        .from('teams')
        .select('*')
        .eq('user_id', currentUserId);

      if (error) throw error;
      setTeams(data || []);
      if (data && data.length > 0) {
        setSelectedTeam(data[0].id);
        fetchTeamMembers(data[0].id);
      }
    } catch (err) {
      setError('Failed to load teams');
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  const fetchTeamMembers = async (teamId: string) => {
    try {
      const { data, error } = await supabase
        .from('team_members')
        .select('*')
        .eq('team_id', teamId);

      if (error) throw error;
      setMembers(data || []);
    } catch (err) {
      console.error('Failed to load team members:', err);
    }
  };

  const handleCreateTeam = async () => {
    if (!newTeamName.trim() || !currentUserId) return;

    try {
      setLoading(true);
      const { data, error } = await supabase
        .from('teams')
        .insert({
          user_id: currentUserId,
          name: newTeamName
        })
        .select()
        .single();

      if (error) throw error;

      // Add current user as owner
      await supabase.from('team_members').insert({
        team_id: data.id,
        user_id: currentUserId,
        role: 'owner'
      });

      setTeams([...teams, data]);
      setSelectedTeam(data.id);
      setNewTeamName('');
      setShowCreateTeam(false);
      setSuccess('Team created successfully');
      setTimeout(() => setSuccess(''), 3000);
    } catch (err) {
      setError('Failed to create team');
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  const handleAddMember = async () => {
    if (!newMemberEmail.trim() || !selectedTeam) return;

    try {
      setLoading(true);

      // First, we need to find the user by email
      // In a real app, you might have a separate endpoint for this
      // For now, we'll create a temporary user entry or use an invite system

      // Add team member (this would normally be done after user accepts invite)
      const { error } = await supabase.from('team_members').insert({
        team_id: selectedTeam,
        user_id: null, // Would be set when user accepts invite
        role: 'member'
      });

      if (error && !error.message.includes('violates unique constraint')) {
        throw error;
      }

      setNewMemberEmail('');
      setShowAddMember(false);
      setSuccess('Team member added successfully');
      setTimeout(() => setSuccess(''), 3000);
      if (selectedTeam) {
        fetchTeamMembers(selectedTeam);
      }
    } catch (err) {
      setError('Failed to add team member');
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  const handleRemoveMember = async (memberId: string) => {
    try {
      setLoading(true);
      const { error } = await supabase
        .from('team_members')
        .delete()
        .eq('id', memberId);

      if (error) throw error;

      setMembers(members.filter(m => m.id !== memberId));
      setSuccess('Team member removed successfully');
      setTimeout(() => setSuccess(''), 3000);
    } catch (err) {
      setError('Failed to remove team member');
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-slate-800 to-slate-900 p-6">
      <div className="max-w-6xl mx-auto">
        {/* Header */}
        <div className="flex items-center justify-between mb-8">
          <div className="flex items-center space-x-3">
            <Users className="h-8 w-8 text-blue-400" />
            <h1 className="text-3xl font-bold text-white">Team Management</h1>
          </div>
          <button
            onClick={() => setShowCreateTeam(true)}
            className="flex items-center space-x-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
          >
            <Plus className="h-5 w-5" />
            <span>New Team</span>
          </button>
        </div>

        {/* Success/Error Messages */}
        {success && (
          <div className="mb-4 p-4 bg-green-900/30 border border-green-700 rounded-lg text-green-400">
            {success}
          </div>
        )}
        {error && (
          <div className="mb-4 p-4 bg-red-900/30 border border-red-700 rounded-lg text-red-400">
            {error}
          </div>
        )}

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          {/* Teams List */}
          <div className="lg:col-span-1">
            <div className="bg-slate-800/60 backdrop-blur-sm border border-slate-700/50 rounded-xl p-6">
              <h2 className="text-lg font-semibold text-white mb-4">Your Teams</h2>
              <div className="space-y-2">
                {teams.length === 0 ? (
                  <p className="text-gray-400 text-sm">No teams yet. Create one to get started!</p>
                ) : (
                  teams.map(team => (
                    <button
                      key={team.id}
                      onClick={() => {
                        setSelectedTeam(team.id);
                        fetchTeamMembers(team.id);
                      }}
                      className={`w-full text-left px-4 py-3 rounded-lg transition-colors ${
                        selectedTeam === team.id
                          ? 'bg-blue-600 text-white'
                          : 'bg-slate-700/50 text-gray-300 hover:bg-slate-700'
                      }`}
                    >
                      <p className="font-medium">{team.name}</p>
                      <p className="text-xs opacity-70">
                        {new Date(team.created_at).toLocaleDateString()}
                      </p>
                    </button>
                  ))
                )}
              </div>
            </div>
          </div>

          {/* Team Members */}
          <div className="lg:col-span-2">
            <div className="bg-slate-800/60 backdrop-blur-sm border border-slate-700/50 rounded-xl p-6">
              <div className="flex items-center justify-between mb-6">
                <h2 className="text-lg font-semibold text-white">Team Members</h2>
                {selectedTeam && (
                  <button
                    onClick={() => setShowAddMember(true)}
                    className="flex items-center space-x-2 px-3 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors text-sm"
                  >
                    <Plus className="h-4 w-4" />
                    <span>Add Member</span>
                  </button>
                )}
              </div>

              {selectedTeam ? (
                <div className="space-y-3">
                  {members.length === 0 ? (
                    <p className="text-gray-400 text-sm">No members yet. Invite your first team member!</p>
                  ) : (
                    members.map(member => (
                      <div
                        key={member.id}
                        className="flex items-center justify-between bg-slate-700/50 p-4 rounded-lg hover:bg-slate-700 transition-colors"
                      >
                        <div className="flex items-center space-x-3">
                          <User className="h-5 w-5 text-gray-400" />
                          <div>
                            <p className="text-white font-medium">{member.name || 'Pending'}</p>
                            <p className="text-gray-400 text-xs">{member.email}</p>
                          </div>
                        </div>
                        <div className="flex items-center space-x-3">
                          <span className={`px-2 py-1 rounded text-xs font-medium ${
                            member.role === 'owner'
                              ? 'bg-purple-600/30 text-purple-300'
                              : 'bg-blue-600/30 text-blue-300'
                          }`}>
                            {member.role.charAt(0).toUpperCase() + member.role.slice(1)}
                          </span>
                          {member.role !== 'owner' && (
                            <button
                              onClick={() => handleRemoveMember(member.id)}
                              className="text-gray-400 hover:text-red-400 transition-colors"
                            >
                              <Trash2 className="h-4 w-4" />
                            </button>
                          )}
                        </div>
                      </div>
                    ))
                  )}
                </div>
              ) : (
                <p className="text-gray-400 text-sm">Select a team to view members</p>
              )}
            </div>
          </div>
        </div>
      </div>

      {/* Create Team Modal */}
      {showCreateTeam && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-50">
          <div className="bg-slate-800 border border-slate-700 rounded-xl p-6 w-full max-w-md">
            <div className="flex items-center justify-between mb-4">
              <h2 className="text-lg font-semibold text-white">Create New Team</h2>
              <button
                onClick={() => setShowCreateTeam(false)}
                className="text-gray-400 hover:text-white transition-colors"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            <div className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-white/80 mb-2">
                  Team Name
                </label>
                <input
                  type="text"
                  value={newTeamName}
                  onChange={(e) => setNewTeamName(e.target.value)}
                  placeholder="e.g., Design Team"
                  className="w-full px-3 py-2 bg-slate-700/50 border border-slate-600 rounded-lg text-white placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-blue-500/50"
                />
              </div>
            </div>

            <div className="flex space-x-3 mt-6">
              <button
                onClick={handleCreateTeam}
                disabled={loading}
                className="flex-1 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors disabled:opacity-50"
              >
                {loading ? 'Creating...' : 'Create Team'}
              </button>
              <button
                onClick={() => setShowCreateTeam(false)}
                className="px-4 py-2 text-gray-400 border border-slate-600 rounded-lg hover:bg-slate-700 transition-colors"
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Add Member Modal */}
      {showAddMember && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-50">
          <div className="bg-slate-800 border border-slate-700 rounded-xl p-6 w-full max-w-md">
            <div className="flex items-center justify-between mb-4">
              <h2 className="text-lg font-semibold text-white">Add Team Member</h2>
              <button
                onClick={() => setShowAddMember(false)}
                className="text-gray-400 hover:text-white transition-colors"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            <p className="text-gray-400 text-sm mb-4">
              Add up to 5 team members to your team. They'll receive an invitation to join.
            </p>

            <div className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-white/80 mb-2">
                  Email Address
                </label>
                <input
                  type="email"
                  value={newMemberEmail}
                  onChange={(e) => setNewMemberEmail(e.target.value)}
                  placeholder="coworker@company.com"
                  className="w-full px-3 py-2 bg-slate-700/50 border border-slate-600 rounded-lg text-white placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-blue-500/50"
                />
              </div>
            </div>

            <div className="flex space-x-3 mt-6">
              <button
                onClick={handleAddMember}
                disabled={loading}
                className="flex-1 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors disabled:opacity-50"
              >
                {loading ? 'Adding...' : 'Add Member'}
              </button>
              <button
                onClick={() => setShowAddMember(false)}
                className="px-4 py-2 text-gray-400 border border-slate-600 rounded-lg hover:bg-slate-700 transition-colors"
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default TeamManagement;
