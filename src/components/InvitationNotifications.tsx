import React, { useState, useEffect } from 'react';
import { Mail, X, Check } from 'lucide-react';

interface Invitation {
  id: string;
  teamId: string;
  teamName: string;
  memberEmail: string;
  memberName: string;
  status: 'pending' | 'accepted' | 'rejected';
  createdAt: string;
}

const InvitationNotifications: React.FC = () => {
  const [invitations, setInvitations] = useState<Invitation[]>([]);
  const [showPanel, setShowPanel] = useState(false);

  useEffect(() => {
    loadInvitations();
  }, []);

  const loadInvitations = () => {
    try {
      const stored = localStorage.getItem('invitations');
      if (stored) {
        const allInvitations = JSON.parse(stored);
        const pending = allInvitations.filter(
          (inv: Invitation) => inv.status === 'pending'
        );
        setInvitations(pending);
      }
    } catch (err) {
      console.error('Failed to load invitations:', err);
    }
  };

  const handleAccept = (invitationId: string) => {
    const stored = JSON.parse(localStorage.getItem('invitations') || '[]');
    const updated = stored.map((inv: Invitation) =>
      inv.id === invitationId
        ? { ...inv, status: 'accepted' }
        : inv
    );
    localStorage.setItem('invitations', JSON.stringify(updated));
    setInvitations(invitations.filter(inv => inv.id !== invitationId));
  };

  const handleReject = (invitationId: string) => {
    const stored = JSON.parse(localStorage.getItem('invitations') || '[]');
    const updated = stored.map((inv: Invitation) =>
      inv.id === invitationId
        ? { ...inv, status: 'rejected' }
        : inv
    );
    localStorage.setItem('invitations', JSON.stringify(updated));
    setInvitations(invitations.filter(inv => inv.id !== invitationId));
  };

  if (invitations.length === 0) {
    return null;
  }

  return (
    <>
      {/* Notification Badge */}
      <button
        onClick={() => setShowPanel(!showPanel)}
        className="fixed bottom-8 left-8 flex items-center justify-center w-12 h-12 bg-gradient-to-r from-amber-500 to-orange-600 rounded-full shadow-lg hover:shadow-xl transition-all hover:scale-110 relative z-40"
        title={`${invitations.length} pending invitation${invitations.length > 1 ? 's' : ''}`}
      >
        <Mail className="h-6 w-6 text-white" />
        {invitations.length > 0 && (
          <span className="absolute -top-2 -right-2 bg-red-500 text-white text-xs font-bold w-5 h-5 rounded-full flex items-center justify-center">
            {invitations.length}
          </span>
        )}
      </button>

      {/* Notifications Panel */}
      {showPanel && (
        <div className="fixed bottom-24 left-8 w-80 bg-slate-800 border border-slate-700 rounded-xl shadow-2xl z-50 max-h-96 overflow-y-auto">
          <div className="sticky top-0 bg-slate-800 border-b border-slate-700 px-4 py-3 flex items-center justify-between">
            <h3 className="text-white font-semibold flex items-center space-x-2">
              <Mail className="h-5 w-5 text-amber-400" />
              <span>Invitations ({invitations.length})</span>
            </h3>
            <button
              onClick={() => setShowPanel(false)}
              className="text-gray-400 hover:text-white transition-colors"
            >
              <X className="h-5 w-5" />
            </button>
          </div>

          <div className="divide-y divide-slate-700">
            {invitations.map((invitation) => (
              <div
                key={invitation.id}
                className="p-4 hover:bg-slate-700/30 transition-colors"
              >
                <div className="mb-2">
                  <p className="text-white font-semibold text-sm">
                    {invitation.teamName}
                  </p>
                  <p className="text-gray-400 text-xs">
                    Invited by {invitation.memberName}
                  </p>
                </div>

                <div className="flex gap-2 mt-3">
                  <button
                    onClick={() => handleAccept(invitation.id)}
                    className="flex-1 flex items-center justify-center space-x-1 px-3 py-2 bg-green-600 hover:bg-green-700 text-white rounded text-sm font-medium transition-colors"
                  >
                    <Check className="h-4 w-4" />
                    <span>Accept</span>
                  </button>
                  <button
                    onClick={() => handleReject(invitation.id)}
                    className="flex-1 flex items-center justify-center space-x-1 px-3 py-2 bg-red-600 hover:bg-red-700 text-white rounded text-sm font-medium transition-colors"
                  >
                    <X className="h-4 w-4" />
                    <span>Decline</span>
                  </button>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
    </>
  );
};

export default InvitationNotifications;
