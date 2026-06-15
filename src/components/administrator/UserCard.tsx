import React from 'react';
import Image from 'next/image';
import { User } from '@/lib/types';
//import defaultProfile from 'defaultProfile.jpg';

interface UserCardProps {
  user: User;
  onClick: () => void;
}

const UserCard: React.FC<UserCardProps> = ({ user, onClick }) => {
  return (
    <div
      onClick={onClick}
      className="bg-gray-50 rounded-lg shadow-md p-4 flex flex-col items-center text-center cursor-pointer hover:shadow-xl hover:scale-105 transition-transform duration-200"
    >
      <h3 className="text-xl font-semibold text-gray-800">{user.username || 'No Name'}</h3>
      <p className="text-gray-500">{user.email || 'No Email'}</p>
      <p className="text-gray-500">{user.role || 'No Role'}</p>
    </div>
  );
};

export default UserCard;