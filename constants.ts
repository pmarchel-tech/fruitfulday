
import { TaskStatus } from './types';

export const STATUS_COLORS: Record<TaskStatus, string> = {
  [TaskStatus.NOT_YET]: 'bg-stone-200 text-stone-600 border-stone-300',
  [TaskStatus.PROGRESS]: 'bg-blue-100 text-blue-800 border-blue-200',
  [TaskStatus.DONE]: 'bg-emerald-100 text-emerald-800 border-emerald-200',
  [TaskStatus.CANCEL]: 'bg-red-100 text-red-800 border-red-200',
  [TaskStatus.PENDING]: 'bg-amber-100 text-amber-800 border-amber-200',
  [TaskStatus.REPETITIVE]: 'bg-purple-100 text-purple-800 border-purple-200',
  [TaskStatus.FOLLOW_UP]: 'bg-orange-100 text-orange-800 border-orange-200',
};

export const CATEGORIES = ['ACSENT', 'PERSONAL', 'QAD', 'UID'];
export const STATUSES = Object.values(TaskStatus);

export const INDONESIAN_HOLIDAYS: Record<string, string> = {
  // 2024
  '2024-01-01': 'Tahun Baru Masehi',
  '2024-02-08': 'Isra Mikraj Nabi Muhammad SAW',
  '2024-02-10': 'Tahun Baru Imlek 2575 Kongzili',
  '2024-03-11': 'Hari Suci Nyepi',
  '2024-03-29': 'Wafat Yesus Kristus',
  '2024-03-31': 'Hari Paskah',
  '2024-04-10': 'Hari Raya Idul Fitri 1445 H',
  '2024-04-11': 'Hari Raya Idul Fitri 1445 H',
  '2024-05-01': 'Hari Buruh Internasional',
  '2024-05-09': 'Kenaikan Yesus Kristus',
  '2024-05-23': 'Hari Raya Waisak 2568 BE',
  '2024-06-01': 'Hari Lahir Pancasila',
  '2024-06-17': 'Hari Raya Idul Adha 1445 H',
  '2024-07-07': 'Tahun Baru Islam 1446 H',
  '2024-08-17': 'Hari Kemerdekaan RI',
  '2024-09-16': 'Maulid Nabi Muhammad SAW',
  '2024-12-25': 'Hari Raya Natal',
  '2024-12-26': 'Cuti Bersama Natal',

  // 2025
  '2025-01-01': 'Tahun Baru Masehi',
  '2025-01-27': 'Isra Mikraj Nabi Muhammad SAW',
  '2025-01-29': 'Tahun Baru Imlek 2576 Kongzili',
  '2025-03-29': 'Hari Suci Nyepi',
  '2025-03-31': 'Hari Raya Idul Fitri 1446 H',
  '2025-04-01': 'Hari Raya Idul Fitri 1446 H',
  '2025-04-18': 'Wafat Yesus Kristus',
  '2025-05-01': 'Hari Buruh Internasional',
  '2025-05-12': 'Hari Raya Waisak 2569 BE',
  '2025-05-29': 'Kenaikan Yesus Kristus',
  '2025-06-01': 'Hari Lahir Pancasila',
  '2025-06-06': 'Hari Raya Idul Adha 1446 H',
  '2025-06-27': 'Tahun Baru Islam 1447 H',
  '2025-08-17': 'Hari Kemerdekaan RI',
  '2025-09-05': 'Maulid Nabi Muhammad SAW',
  '2025-12-25': 'Hari Raya Natal',
  '2025-12-26': 'Cuti Bersama Natal',

  // 2026
  '2026-01-01': 'Tahun Baru Masehi',
  '2026-01-15': 'Isra Mikraj Nabi Muhammad SAW',
  '2026-02-17': 'Tahun Baru Imlek 2577 Kongzili',
  '2026-03-19': 'Hari Suci Nyepi',
  '2026-03-20': 'Hari Raya Idul Fitri 1447 H',
  '2026-03-21': 'Hari Raya Idul Fitri 1447 H',
  '2026-04-03': 'Wafat Yesus Kristus',
  '2026-05-01': 'Hari Buruh Internasional',
  '2026-05-14': 'Kenaikan Yesus Kristus',
  '2026-05-27': 'Hari Raya Idul Adha 1447 H',
  '2026-06-01': 'Hari Lahir Pancasila',
  '2026-06-16': 'Tahun Baru Islam 1448 H',
  '2026-08-17': 'Hari Kemerdekaan RI',
  '2026-08-25': 'Maulid Nabi Muhammad SAW',
  '2026-12-25': 'Hari Raya Natal',
  '2026-12-26': 'Cuti Bersama Natal',

  // 2027
  '2027-01-01': 'Tahun Baru Masehi',
  '2027-01-05': 'Isra Mikraj Nabi Muhammad SAW',
  '2027-02-06': 'Tahun Baru Imlek 2578 Kongzili',
  '2027-03-09': 'Hari Raya Idul Fitri 1448 H',
  '2027-03-10': 'Hari Raya Idul Fitri 1448 H',
  '2027-03-26': 'Wafat Yesus Kristus',
  '2027-05-01': 'Hari Buruh Internasional',
  '2027-05-06': 'Kenaikan Yesus Kristus',
  '2027-05-16': 'Hari Raya Idul Adha 1448 H',
  '2027-05-20': 'Hari Raya Waisak 2571 BE',
  '2027-06-01': 'Hari Lahir Pancasila',
  '2027-06-06': 'Tahun Baru Islam 1449 H',
  '2027-08-17': 'Hari Kemerdekaan RI',
  '2027-08-15': 'Maulid Nabi Muhammad SAW',
  '2027-12-25': 'Hari Raya Natal',
  '2027-12-26': 'Cuti Bersama Natal',
};