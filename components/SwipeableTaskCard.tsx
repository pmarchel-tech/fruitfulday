import React, { useState, useRef } from 'react';
import { Task, TaskStatus, User, TaskUpdate } from '../types';
import { Check, RefreshCw, Calendar } from 'lucide-react';

interface SwipeableTaskCardProps {
  task: Task;
  isSelected: boolean;
  onSelect: () => void;
  onToggleStatus: (task: Task) => void;
  currentUser: User;
  userMap: { [id: string]: User };
  updates: TaskUpdate[];
}

const formatDDMMM = (dateStr: string) => {
  if (!dateStr) return '';
  try {
    const [year, month, day] = dateStr.split('-').map(Number);
    const date = new Date(year, month - 1, day);
    if (isNaN(date.getTime())) return dateStr;
    const dd = day.toString().padStart(2, '0');
    const mmm = date.toLocaleDateString('en-US', { month: 'short' });
    return `${dd}-${mmm}`;
  } catch (e) {
    return dateStr;
  }
};

const getDaysRemaining = (targetDate: string) => {
  if (!targetDate) return 0;
  try {
    const target = new Date(targetDate);
    const today = new Date();
    target.setHours(0, 0, 0, 0);
    today.setHours(0, 0, 0, 0);
    const diffTime = target.getTime() - today.getTime();
    return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  } catch (e) {
    return 0;
  }
};

const SwipeableTaskCard: React.FC<SwipeableTaskCardProps> = ({
  task,
  isSelected,
  onSelect,
  onToggleStatus,
  currentUser,
  userMap,
  updates
}) => {
  const [offsetX, setOffsetX] = useState(0);
  const [isDragging, setIsDragging] = useState(false);
  const startX = useRef(0);
  const startY = useRef(0);

  const isCompleted = task.status === TaskStatus.DONE;
  const daysRemaining = getDaysRemaining(task.targetDate);
  const isOverdue = daysRemaining < 0 && !isCompleted;
  const taskActionsCount = updates.filter(u => u.taskId === task.id).length;

  const handleStart = (clientX: number, clientY: number) => {
    setIsDragging(true);
    startX.current = clientX;
    startY.current = clientY;
  };

  const handleMove = (clientX: number, clientY: number) => {
    if (!isDragging) return;
    const diffX = clientX - startX.current;
    const diffY = clientY - startY.current;

    // Scroll guard: if vertical movement is greater than horizontal, cancel dragging
    // to let the browser scroll naturally.
    if (Math.abs(diffY) > Math.abs(diffX) && Math.abs(diffY) > 5) {
      setIsDragging(false);
      setOffsetX(0);
      return;
    }

    // Only allow swiping to the right (positive offsetX)
    if (diffX > 0) {
      // Limit swipe range to a max of 160px for visual comfort
      setOffsetX(Math.min(diffX, 160));
    } else {
      setOffsetX(0);
    }
  };

  const handleEnd = () => {
    setIsDragging(false);
    const dragDistance = offsetX;
    
    if (dragDistance > 90) {
      // Swiped far enough to the right: trigger toggle status
      onToggleStatus(task);
    } else if (dragDistance < 6) {
      // Tap or click: select the task
      onSelect();
    }
    setOffsetX(0);
  };

  // Mouse event handlers
  const onMouseDown = (e: React.MouseEvent) => handleStart(e.clientX, e.clientY);
  const onMouseMove = (e: React.MouseEvent) => handleMove(e.clientX, e.clientY);
  const onMouseUp = () => handleEnd();
  const onMouseLeave = () => { if (isDragging) handleEnd(); };

  // Touch event handlers for mobile
  const onTouchStart = (e: React.TouchEvent) => handleStart(e.touches[0].clientX, e.touches[0].clientY);
  const onTouchMove = (e: React.TouchEvent) => handleMove(e.touches[0].clientX, e.touches[0].clientY);
  const onTouchEnd = () => handleEnd();

  const getBackgroundClass = () => {
    if (offsetX > 40) {
      return isCompleted ? 'bg-indigo-500 justify-start' : 'bg-emerald-500 justify-start';
    }
    return 'bg-slate-100';
  };

  return (
    <div className="relative w-full select-none group touch-pan-y">
      {/* Background Actions Layer (Shown when swiped) */}
      <div className={`absolute inset-0 flex items-center px-8 text-white font-bold transition-all rounded-[24px] ${getBackgroundClass()}`}>
        {offsetX > 40 && (
          <div className="flex items-center gap-2.5 animate-fade-in">
            {isCompleted ? (
              <>
                <RefreshCw size={18} className="animate-spin-slow" />
                <span className="text-sm tracking-wide">Mark Incomplete</span>
              </>
            ) : (
              <>
                <Check size={20} className="stroke-[3.5]" />
                <span className="text-sm tracking-wide">Mark Complete</span>
              </>
            )}
          </div>
        )}
      </div>

      {/* Foreground Task Content Card Layer */}
      <div
        className={`relative bg-white p-5 rounded-[24px] flex items-center justify-between border-2 transition-all cursor-grab active:cursor-grabbing z-10 ${
          isSelected
            ? 'border-[#0038FF] ring-4 ring-[#0038FF]/5 shadow-md shadow-blue-500/5'
            : 'border-slate-100/80 hover:border-slate-200 hover:shadow-sm'
        }`}
        style={{ transform: `translateX(${offsetX}px)`, transition: isDragging ? 'none' : 'transform 0.2s ease-out' }}
        onMouseDown={onMouseDown}
        onMouseMove={onMouseMove}
        onMouseUp={onMouseUp}
        onMouseLeave={onMouseLeave}
        onTouchStart={onTouchStart}
        onTouchMove={onTouchMove}
        onTouchEnd={onTouchEnd}
      >
        {/* Left details */}
        <div className="flex-1 min-w-0 pr-4 pointer-events-none">
          <h3 className={`font-bold text-base leading-tight truncate transition-all duration-300 ${
            isCompleted ? 'text-[#8E9BB2] line-through font-semibold' : 'text-[#1C2038]'
          }`}>
            {task.title}
          </h3>
          
          <div className="flex flex-wrap items-center gap-1.5 mt-2">
            <span className="inline-flex items-center px-2 py-0.5 rounded-lg text-[8.5px] font-bold bg-[#E6ECFF] text-[#0038FF] uppercase tracking-wider">
              {task.category}
            </span>
            
            {isOverdue ? (
              <span className="inline-flex items-center px-2 py-0.5 rounded-lg text-[8.5px] font-bold bg-[#FFEBEC] text-[#D83F52] uppercase tracking-wider">
                Overdue
              </span>
            ) : isCompleted ? (
              <span className="inline-flex items-center px-2 py-0.5 rounded-lg text-[8.5px] font-bold bg-[#E6F4EA] text-[#137333] uppercase tracking-wider">
                Done
              </span>
            ) : (
              <span className="inline-flex items-center px-2 py-0.5 rounded-lg text-[8.5px] font-bold bg-slate-100 text-[#5F6368] uppercase tracking-wider">
                In Progress
              </span>
            )}

            <span className="inline-flex items-center px-2 py-0.5 rounded-lg text-[8.5px] font-bold bg-slate-100 text-slate-600 uppercase tracking-wider" title="Due Date">
              {formatDDMMM(task.targetDate)}
            </span>

            <span className="inline-flex items-center px-2 py-0.5 rounded-lg text-[8.5px] font-bold bg-purple-50 text-purple-700 border border-purple-100/60 uppercase tracking-wider" title="Actions Count">
              {taskActionsCount}
            </span>

            {task.tags?.map(t => (
              <span key={t} className="inline-flex items-center px-1.5 py-0.5 rounded-md text-[8.5px] font-extrabold bg-blue-50/70 text-[#0038FF] border border-blue-100/30 uppercase tracking-tight">
                #{t}
              </span>
            ))}

            {/* Show owner small badge for admin */}
            {currentUser.role === 'ADMIN' && task.userId !== currentUser.id && userMap[task.userId] && (
              <span className="inline-flex items-center px-2 py-0.5 rounded-lg text-[8.5px] font-bold bg-amber-50 text-amber-700 border border-amber-100" title="Task Owner">
                {userMap[task.userId].username}
              </span>
            )}
          </div>
        </div>

        {/* Right Indicator (Visual cue for swiping, replacing the check button) */}
        <div className="flex items-center gap-1 text-slate-300 pointer-events-none shrink-0">
          <ChevronRightIcon className="text-slate-350 opacity-60" />
        </div>
      </div>
    </div>
  );
};

// Small Chevron icon component to avoid importing another library
const ChevronRightIcon = ({ className }: { className?: string }) => (
  <svg 
    xmlns="http://www.w3.org/2000/svg" 
    viewBox="0 0 24 24" 
    fill="none" 
    stroke="currentColor" 
    strokeWidth="2.5" 
    strokeLinecap="round" 
    strokeLinejoin="round" 
    className={`w-5 h-5 ${className}`}
  >
    <path d="m9 18 6-6-6-6" />
  </svg>
);

export default SwipeableTaskCard;
