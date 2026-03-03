'use client';
import { useAuth } from '@/contexts/AuthContext';

export default function ProfilePage() {
  const { user, loading } = useAuth();

  if (loading) return <p>Loading...</p>;
  if (!user) return <p>Please log in.</p>;

  console.log("User Data:", user);

  return (
    <div>
      <h1>Welcome, {user.username}</h1>
      <p>Your Role: {user.role}</p>
      <p>Email: {user.email}</p>
      <p>{user.uid}</p>
    </div>
  );
}