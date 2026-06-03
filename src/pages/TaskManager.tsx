import React, { useState, useEffect } from 'react';
import { DragDropContext, Droppable, Draggable, DropResult } from 'react-beautiful-dnd';
import { Plus, Search, Filter, Calendar, User, MessageSquare, Paperclip, MoreHorizontal, Star, Users } from 'lucide-react';
import { useLanguage } from '../contexts/LanguageContext';
import { supabase } from '../services/supabaseClient';

interface Task {
  id: string;
  title: string;
  description: string;
  assignee: string;
  assigneeId: string;
  avatar?: string;
  dueDate: string;
  priority: 'low' | 'medium' | 'high';
  status: 'backlog' | 'todo' | 'inProgress' | 'review' | 'done';
  comments: number;
  attachments: number;
  tags: string[];
  color: string;
}

interface TeamMember {
  id: string;
  name: string;
  email: string;
}

const TaskManager: React.FC = () => {
  const { t } = useLanguage();
  const [showAddTask, setShowAddTask] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [filterType, setFilterType] = useState('all');
  const [filterAssignee, setFilterAssignee] = useState('all');
  const [filterPriority, setFilterPriority] = useState('all');
  const [filterStage, setFilterStage] = useState('all');
  const [teamMembers, setTeamMembers] = useState<TeamMember[]>([]);
  const [loading, setLoading] = useState(false);
  const [selectedTeam, setSelectedTeam] = useState<string | null>(null);
  const [newTask, setNewTask] = useState({
    title: '',
    description: '',
    assigneeId: '',
    dueDate: '',
    priority: 'medium' as const
  });

  const [tasks, setTasks] = useState<Task[]>([
    {
      id: '1',
      title: 'Create Blog Content',
      description: 'Write a blog about our design process and publish it on our website',
      assignee: 'John Doe',
      assigneeId: '1',
      avatar: 'JD',
      dueDate: '2024-01-20',
      priority: 'high',
      status: 'backlog',
      comments: 3,
      attachments: 2,
      tags: ['Content', 'Marketing'],
      color: 'border-l-green-500'
    },
    {
      id: '2',
      title: 'Send Email to John',
      description: 'Make sure to talk about what the design process and get his approval',
      assignee: 'Sarah Smith',
      assigneeId: '2',
      avatar: 'SS',
      dueDate: '2024-01-25',
      priority: 'medium',
      status: 'backlog',
      comments: 2,
      attachments: 1,
      tags: ['Email', 'Communication'],
      color: 'border-l-orange-500'
    },
    {
      id: '3',
      title: 'Review UX Change',
      description: 'Look through the document with the details of UX changes and other things',
      assignee: 'Mike Johnson',
      assigneeId: '3',
      avatar: 'MJ',
      dueDate: '2024-01-15',
      priority: 'low',
      status: 'backlog',
      comments: 4,
      attachments: 0,
      tags: ['UX', 'Review'],
      color: 'border-l-red-500'
    },
    {
      id: '4',
      title: 'Send Out Proposals',
      description: 'Find a new client and send them a detailed proposal for our services',
      assignee: 'Emily Brown',
      assigneeId: '4',
      avatar: 'EB',
      dueDate: '2024-01-18',
      priority: 'high',
      status: 'todo',
      comments: 1,
      attachments: 3,
      tags: ['Sales', 'Proposal'],
      color: 'border-l-blue-500'
    },
    {
      id: '5',
      title: 'Design the Dashboard',
      description: 'Be sure to look at what\'s been done and implement the design',
      assignee: 'Alex Wilson',
      assigneeId: '5',
      avatar: 'AW',
      dueDate: '2024-01-22',
      priority: 'medium',
      status: 'todo',
      comments: 2,
      attachments: 1,
      tags: ['Design', 'UI'],
      color: 'border-l-purple-500'
    }
  ]);

  const columns = [
    {
      id: 'backlog',
      title: 'BACKLOG',
      color: 'text-gray-400',
      bgColor: 'bg-slate-800/50'
    },
    {
      id: 'todo',
      title: 'TO DO',
      color: 'text-blue-400',
      bgColor: 'bg-blue-900/20'
    },
    {
      id: 'inProgress',
      title: 'IN PROGRESS',
      color: 'text-yellow-400',
      bgColor: 'bg-yellow-900/20'
    },
    {
      id: 'review',
      title: 'REVIEW',
      color: 'text-purple-400',
      bgColor: 'bg-purple-900/20'
    },
    {
      id: 'done',
      title: 'DONE',
      color: 'text-green-400',
      bgColor: 'bg-green-900/20'
    }
  ];

  useEffect(() => {
    fetchTeamMembers();
  }, []);

  const fetchTeamMembers = async () => {
    try {
      // For now, use mock data
      // In production, fetch from Supabase
      setTeamMembers([
        { id: '1', name: 'John Doe', email: 'john@example.com' },
        { id: '2', name: 'Sarah Smith', email: 'sarah@example.com' },
        { id: '3', name: 'Mike Johnson', email: 'mike@example.com' },
        { id: '4', name: 'Emily Brown', email: 'emily@example.com' },
        { id: '5', name: 'Alex Wilson', email: 'alex@example.com' }
      ]);
    } catch (err) {
      console.error('Failed to load team members:', err);
    }
  };

  const filteredTasks = tasks.filter(task => {
    const matchesSearch = task.title.toLowerCase().includes(searchTerm.toLowerCase()) ||
                         task.description.toLowerCase().includes(searchTerm.toLowerCase());
    const matchesType = filterType === 'all' || true;
    const matchesAssignee = filterAssignee === 'all' || task.assigneeId === filterAssignee;
    const matchesPriority = filterPriority === 'all' || task.priority === filterPriority;
    const matchesStage = filterStage === 'all' || task.status === filterStage;

    return matchesSearch && matchesType && matchesAssignee && matchesPriority && matchesStage;
  });

  const getTasksByStatus = (status: string) => {
    return filteredTasks.filter(task => task.status === status);
  };

  const handleDragEnd = (result: DropResult) => {
    const { destination, source, draggableId } = result;

    if (!destination) {
      return;
    }

    if (
      destination.droppableId === source.droppableId &&
      destination.index === source.index
    ) {
      return;
    }

    const taskToMove = tasks.find(task => task.id === draggableId);
    if (!taskToMove) {
      return;
    }

    const newTasks = tasks.map(task => {
      if (task.id === draggableId) {
        return {
          ...task,
          status: destination.droppableId as Task['status']
        };
      }
      return task;
    });

    setTasks(newTasks);
  };

  const handleAddTask = () => {
    if (!newTask.title) return;

    const assignedMember = teamMembers.find(m => m.id === newTask.assigneeId);
    const task: Task = {
      id: Date.now().toString(),
      title: newTask.title,
      description: newTask.description,
      assignee: assignedMember?.name || 'Unassigned',
      assigneeId: newTask.assigneeId || 'unassigned',
      status: 'backlog',
      comments: 0,
      attachments: 0,
      tags: ['New'],
      color: 'border-l-blue-500',
      avatar: assignedMember?.name.split(' ').map(n => n[0]).join(''),
      dueDate: newTask.dueDate,
      priority: newTask.priority
    };

    setTasks(prev => [...prev, task]);
    setNewTask({
      title: '',
      description: '',
      assigneeId: '',
      dueDate: '',
      priority: 'medium'
    });
    setShowAddTask(false);
  };

  const getPriorityColor = (priority: string) => {
    switch (priority) {
      case 'high':
        return 'bg-red-500';
      case 'medium':
        return 'bg-yellow-500';
      case 'low':
        return 'bg-green-500';
      default:
        return 'bg-gray-500';
    }
  };

  const uniqueAssignees = [...new Set(tasks.map(task => task.assigneeId))];

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-slate-800 to-slate-900 p-6">
      {/* Header */}
      <div className="flex items-center justify-between mb-6">
        <div className="flex items-center space-x-6">
          <div className="flex items-center space-x-3">
            <Star className="h-6 w-6 text-blue-400" />
            <span className="text-white/60 text-sm">Saved Filters</span>
          </div>

          <div className="flex items-center space-x-2 text-white/60">
            <div className="w-4 h-4 bg-blue-500 rounded"></div>
            <div className="w-4 h-4 bg-gray-600 rounded"></div>
          </div>
        </div>

        <div className="flex items-center space-x-4">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 h-4 w-4 text-gray-400" />
            <input
              type="text"
              placeholder="Search Tasks"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="pl-10 pr-4 py-2 w-64 bg-slate-800/50 border border-slate-700 rounded-lg text-white placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-blue-500/50 focus:border-transparent backdrop-blur-sm"
            />
          </div>
        </div>
      </div>

      {/* Top Filters Bar */}
      <div className="bg-slate-800/30 backdrop-blur-sm border border-slate-700/50 rounded-xl p-4 mb-6">
        <div className="flex items-center space-x-2 mb-4">
          <Filter className="h-5 w-5 text-white/70" />
          <span className="text-white/70 font-medium">Filter Tasks</span>
        </div>

        <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
          <div>
            <label className="block text-white/60 text-sm font-medium mb-2">Type</label>
            <select
              value={filterType}
              onChange={(e) => setFilterType(e.target.value)}
              className="w-full px-3 py-2 bg-slate-700/50 border border-slate-600 rounded-lg text-white text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/50"
            >
              <option value="all">All Types</option>
              <option value="task">Task</option>
              <option value="bug">Bug</option>
              <option value="feature">Feature</option>
            </select>
          </div>

          <div>
            <label className="block text-white/60 text-sm font-medium mb-2">Assigned To</label>
            <select
              value={filterAssignee}
              onChange={(e) => setFilterAssignee(e.target.value)}
              className="w-full px-3 py-2 bg-slate-700/50 border border-slate-600 rounded-lg text-white text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/50"
            >
              <option value="all">All Users</option>
              {teamMembers.map(member => (
                <option key={member.id} value={member.id}>{member.name}</option>
              ))}
            </select>
          </div>

          <div>
            <label className="block text-white/60 text-sm font-medium mb-2">Priority</label>
            <select
              value={filterPriority}
              onChange={(e) => setFilterPriority(e.target.value)}
              className="w-full px-3 py-2 bg-slate-700/50 border border-slate-600 rounded-lg text-white text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/50"
            >
              <option value="all">All Priorities</option>
              <option value="high">High</option>
              <option value="medium">Medium</option>
              <option value="low">Low</option>
            </select>
          </div>

          <div>
            <label className="block text-white/60 text-sm font-medium mb-2">Stage</label>
            <select
              value={filterStage}
              onChange={(e) => setFilterStage(e.target.value)}
              className="w-full px-3 py-2 bg-slate-700/50 border border-slate-600 rounded-lg text-white text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/50"
            >
              <option value="all">All Stages</option>
              <option value="backlog">Backlog</option>
              <option value="todo">To Do</option>
              <option value="inProgress">In Progress</option>
              <option value="review">Review</option>
              <option value="done">Done</option>
            </select>
          </div>
        </div>
      </div>

      {/* Kanban Board */}
      <DragDropContext onDragEnd={handleDragEnd}>
        <div className="grid grid-cols-5 gap-4">
          {columns.map((column) => (
            <div key={column.id} className="flex flex-col">
              <div className="flex items-center justify-between mb-4">
                <div className="flex items-center space-x-2">
                  <h2 className={`text-xs font-bold tracking-wider ${column.color}`}>
                    {column.title}
                  </h2>
                  <span className="text-xs text-white/40">
                    ({getTasksByStatus(column.id).length})
                  </span>
                  <button className="w-5 h-5 rounded-full border border-gray-600 flex items-center justify-center hover:bg-gray-700 transition-colors">
                    <Plus className="h-3 w-3 text-gray-400" />
                  </button>
                </div>
              </div>

              <Droppable droppableId={column.id}>
                {(provided, snapshot) => (
                  <div
                    ref={provided.innerRef}
                    {...provided.droppableProps}
                    className={`flex-1 space-y-3 min-h-[600px] p-2 rounded-lg transition-colors ${
                      snapshot.isDraggingOver ? column.bgColor : ''
                    }`}
                  >
                    {getTasksByStatus(column.id).map((task, index) => (
                      <Draggable key={task.id} draggableId={task.id} index={index}>
                        {(provided, snapshot) => (
                          <div
                            ref={provided.innerRef}
                            {...provided.draggableProps}
                            {...provided.dragHandleProps}
                            className={`bg-slate-800/60 backdrop-blur-sm border border-slate-700/50 rounded-lg p-4 hover:bg-slate-800/80 transition-all cursor-grab active:cursor-grabbing ${
                              snapshot.isDragging ? 'shadow-2xl ring-2 ring-blue-500/50 rotate-2 z-50' : 'hover:shadow-lg'
                            } ${task.color} border-l-4`}
                            style={{
                              ...provided.draggableProps.style,
                              zIndex: snapshot.isDragging ? 1000 : 'auto'
                            }}
                          >
                            <div className="flex items-start justify-between mb-3">
                              <h3 className="text-white font-medium text-sm leading-tight pr-2">
                                {task.title}
                              </h3>
                              <button className="text-gray-400 hover:text-white transition-colors">
                                <MoreHorizontal className="h-4 w-4" />
                              </button>
                            </div>

                            <p className="text-gray-400 text-xs mb-4 leading-relaxed">
                              {task.description}
                            </p>

                            <div className="flex flex-wrap gap-1 mb-4">
                              {task.tags.map((tag, tagIndex) => (
                                <span
                                  key={tagIndex}
                                  className="px-2 py-1 bg-slate-700/50 text-gray-300 text-xs rounded-md"
                                >
                                  {tag}
                                </span>
                              ))}
                            </div>

                            <div className="flex items-center justify-between">
                              <div className="flex items-center space-x-3">
                                <div className="flex items-center space-x-1 text-gray-400">
                                  <MessageSquare className="h-3 w-3" />
                                  <span className="text-xs">{task.comments}</span>
                                </div>
                                {task.attachments > 0 && (
                                  <div className="flex items-center space-x-1 text-gray-400">
                                    <Paperclip className="h-3 w-3" />
                                    <span className="text-xs">{task.attachments}</span>
                                  </div>
                                )}
                                <div className="flex items-center space-x-1 text-gray-400">
                                  <Calendar className="h-3 w-3" />
                                  <span className="text-xs">{task.dueDate}</span>
                                </div>
                              </div>

                              <div className="flex items-center space-x-2">
                                <div className={`w-2 h-2 rounded-full ${getPriorityColor(task.priority)}`}></div>
                                <div className="w-6 h-6 bg-gradient-to-r from-blue-500 to-purple-600 rounded-full flex items-center justify-center text-white text-xs font-medium">
                                  {task.avatar}
                                </div>
                              </div>
                            </div>
                          </div>
                        )}
                      </Draggable>
                    ))}
                    {provided.placeholder}
                  </div>
                )}
              </Droppable>
            </div>
          ))}
        </div>
      </DragDropContext>

      {/* Add Task Modal */}
      {showAddTask && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-50">
          <div className="bg-slate-800 border border-slate-700 rounded-xl p-6 w-full max-w-md">
            <h2 className="text-lg font-semibold text-white mb-4">New Task</h2>

            <div className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-white/80 mb-1">
                  Title
                </label>
                <input
                  type="text"
                  value={newTask.title}
                  onChange={(e) => setNewTask(prev => ({ ...prev, title: e.target.value }))}
                  className="w-full px-3 py-2 bg-slate-700/50 border border-slate-600 rounded-lg text-white focus:outline-none focus:ring-2 focus:ring-blue-500/50 focus:border-transparent"
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-white/80 mb-1">
                  Description
                </label>
                <textarea
                  value={newTask.description}
                  onChange={(e) => setNewTask(prev => ({ ...prev, description: e.target.value }))}
                  rows={3}
                  className="w-full px-3 py-2 bg-slate-700/50 border border-slate-600 rounded-lg text-white focus:outline-none focus:ring-2 focus:ring-blue-500/50 focus:border-transparent resize-none"
                />
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-sm font-medium text-white/80 mb-1">
                    Assign To
                  </label>
                  <select
                    value={newTask.assigneeId}
                    onChange={(e) => setNewTask(prev => ({ ...prev, assigneeId: e.target.value }))}
                    className="w-full px-3 py-2 bg-slate-700/50 border border-slate-600 rounded-lg text-white focus:outline-none focus:ring-2 focus:ring-blue-500/50 focus:border-transparent"
                  >
                    <option value="">Select Member</option>
                    {teamMembers.map(member => (
                      <option key={member.id} value={member.id}>{member.name}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-sm font-medium text-white/80 mb-1">
                    Due Date
                  </label>
                  <input
                    type="date"
                    value={newTask.dueDate}
                    onChange={(e) => setNewTask(prev => ({ ...prev, dueDate: e.target.value }))}
                    className="w-full px-3 py-2 bg-slate-700/50 border border-slate-600 rounded-lg text-white focus:outline-none focus:ring-2 focus:ring-blue-500/50 focus:border-transparent"
                  />
                </div>
              </div>

              <div>
                <label className="block text-sm font-medium text-white/80 mb-1">
                  Priority
                </label>
                <select
                  value={newTask.priority}
                  onChange={(e) => setNewTask(prev => ({ ...prev, priority: e.target.value as Task['priority'] }))}
                  className="w-full px-3 py-2 bg-slate-700/50 border border-slate-600 rounded-lg text-white focus:outline-none focus:ring-2 focus:ring-blue-500/50 focus:border-transparent"
                >
                  <option value="low">Low</option>
                  <option value="medium">Medium</option>
                  <option value="high">High</option>
                </select>
              </div>
            </div>

            <div className="flex space-x-3 mt-6">
              <button
                onClick={handleAddTask}
                className="flex-1 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
              >
                Add Task
              </button>
              <button
                onClick={() => setShowAddTask(false)}
                className="px-4 py-2 text-gray-400 border border-slate-600 rounded-lg hover:bg-slate-700 transition-colors"
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Floating Add Button */}
      <button
        onClick={() => setShowAddTask(true)}
        className="fixed bottom-8 right-8 w-14 h-14 bg-gradient-to-r from-blue-600 to-purple-600 rounded-full flex items-center justify-center shadow-lg hover:shadow-xl transition-all duration-200 hover:scale-110"
      >
        <Plus className="h-6 w-6 text-white" />
      </button>
    </div>
  );
};

export default TaskManager;
