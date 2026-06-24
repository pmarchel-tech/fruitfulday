
import React, { useState, useRef } from 'react';
import { TaskUpdate } from '../types';
import { Edit2, Calendar, Trash2, ArrowRight, RefreshCw } from 'lucide-react';

interface SwipeableCardProps {
  update: TaskUpdate;
  onEdit: (update: TaskUpdate) => void;
  onCalendar: (update: TaskUpdate) => void;
  onDelete: (update: TaskUpdate) => void;
}

const SwipeableCard: React.FC<SwipeableCardProps> = ({ update, onEdit, onCalendar, onDelete }) => {
  const [offsetX, setOffsetX] = useState(0);
  const [isDragging, setIsDragging] = useState(false);
  const startX = useRef(0);
  const startY = useRef(0);

  const handleStart = (clientX: number, clientY: number) => {
    setIsDragging(true);
    startX.current = clientX;
    startY.current = clientY;
  };

  const handleMove = (clientX: number, clientY: number) => {
    if (!isDragging) return;
    const diffX = clientX - startX.current;
    const diffY = clientY - startY.current;

    // Check if the user is scrolling vertically
    // If vertical movement is greater than horizontal, let the browser handle scrolling
    if (Math.abs(diffY) > Math.abs(diffX) && Math.abs(diffY) > 5) {
      setIsDragging(false);
      setOffsetX(0);
      return;
    }

    // Limit swipe range
    if (Math.abs(diffX) < 150) {
      setOffsetX(diffX);
    }
  };

  const handleEnd = () => {
    setIsDragging(false);
    if (offsetX > 80) { // Reduced threshold
      onCalendar(update); // Swipe Right -> Calendar
    } else if (offsetX < -80) { // Reduced threshold
      onDelete(update); // Swipe Left -> Delete
    }
    setOffsetX(0);
  };

  // Mouse events
  const onMouseDown = (e: React.MouseEvent) => handleStart(e.clientX, e.clientY);
  const onMouseMove = (e: React.MouseEvent) => handleMove(e.clientX, e.clientY);
  const onMouseUp = () => handleEnd();
  const onMouseLeave = () => { if(isDragging) handleEnd(); };

  // Touch events
  const onTouchStart = (e: React.TouchEvent) => handleStart(e.touches[0].clientX, e.touches[0].clientY);
  const onTouchMove = (e: React.TouchEvent) => handleMove(e.touches[0].clientX, e.touches[0].clientY);
  const onTouchEnd = () => handleEnd();

  const getBackground = () => {
    if (offsetX > 40) return 'bg-emerald-500 justify-start'; // Calendar (Right Swipe)
    if (offsetX < -40) return 'bg-red-500 justify-end'; // Delete (Left Swipe)
    return 'bg-stone-100';
  };

  const formatDate = (isoString: string) => {
    if (!isoString) return '';
    try {
      const [year, month, day] = isoString.split('-').map(Number);
      const date = new Date(year, month - 1, day);
      
      const ddd = date.toLocaleDateString('en-US', { weekday: 'short' });
      const dd = day.toString().padStart(2, '0');
      const mmm = date.toLocaleDateString('en-US', { month: 'short' });
      const yy = year.toString().slice(-2);
      
      return `${ddd}, ${dd} - ${mmm} - ${yy}`;
    } catch (e) {
      return isoString;
    }
  };

  const formatStatus = (s: string) => s.replace(/_/g, ' ');

  return (
    <div className="relative w-full mb-4 select-none group min-h-[100px]">
      {/* Background Actions Layer */}
      <div className={`absolute inset-0 flex items-center px-6 text-white font-bold transition-colors rounded-xl ${getBackground()}`}>
        {offsetX > 0 && <div className="flex items-center gap-2"><Calendar size={20}/> <span>Calendar</span></div>}
        {offsetX < 0 && <div className="flex items-center gap-2"><span>Delete</span> <Trash2 size={20}/></div>}
      </div>

      {/* Foreground Content Layer */}
      <div
        className="relative bg-white shadow-sm border border-stone-200 p-4 rounded-xl cursor-grab active:cursor-grabbing flex flex-col justify-between transition-transform duration-100 ease-linear z-10 min-h-[100px] touch-pan-y"
        style={{ transform: `translateX(${offsetX}px)` }}
        onMouseDown={onMouseDown}
        onMouseMove={onMouseMove}
        onMouseUp={onMouseUp}
        onMouseLeave={onMouseLeave}
        onTouchStart={onTouchStart}
        onTouchMove={onTouchMove}
        onTouchEnd={onTouchEnd}
      >
        <div className="flex justify-between items-start mb-2">
          <div className="flex flex-col gap-1 flex-1 min-w-0">
             <div className="flex justify-between items-center w-full">
                <span className="text-xs font-medium text-stone-400">{formatDate(update.date)}</span>
             </div>
             
             {update.statusChange && (
               <div className="flex items-center gap-3 text-xs font-bold bg-blue-50 border border-blue-200 text-blue-800 px-3 py-2 rounded-lg w-full mt-2 shadow-sm relative overflow-hidden">
                 <div className="absolute left-0 top-0 bottom-0 w-1 bg-blue-400"></div>
                 <RefreshCw size={14} className="shrink-0 text-blue-500" />
                 <div className="flex flex-wrap items-center gap-1.5">
                    <span className="uppercase opacity-60 line-through decoration-blue-400/50">{formatStatus(update.statusChange.from)}</span>
                    <ArrowRight size={14} className="shrink-0 text-blue-400" />
                    <span className="uppercase text-blue-700">{formatStatus(update.statusChange.to)}</span>
                 </div>
               </div>
             )}
          </div>
          
          <div className="flex gap-2 shrink-0 ml-2">
             <button onClick={(e) => { e.stopPropagation(); onEdit(update); }} className="p-1.5 hover:bg-stone-100 rounded-full text-stone-400 hover:text-stone-600 transition-colors" title="Edit">
                <Edit2 size={14} />
             </button>
             <button onClick={(e) => { e.stopPropagation(); onCalendar(update); }} className="p-1.5 hover:bg-emerald-50 rounded-full text-stone-400 hover:text-emerald-600 transition-colors" title="Add to Calendar">
                <Calendar size={14} />
             </button>
             <button onClick={(e) => { e.stopPropagation(); onDelete(update); }} className="p-1.5 hover:bg-red-50 rounded-full text-stone-400 hover:text-red-600 transition-colors" title="Delete">
                <Trash2 size={14} />
             </button>
          </div>
        </div>
        
        <p className="text-sm text-secondary leading-relaxed whitespace-pre-wrap break-words mt-1">
          {update.content}
        </p>
        
        <div className="text-[10px] text-stone-300 text-right mt-2 opacity-50 md:block hidden">
          {offsetX > 0 ? 'Release for Calendar' : offsetX < 0 ? 'Release to Delete' : 'Swipe left delete, right calendar'}
        </div>
      </div>
    </div>
  );
};

export default SwipeableCard;
