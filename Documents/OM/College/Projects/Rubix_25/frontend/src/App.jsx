import React from 'react';
import MessageForm from './components/MessageForm';
import MessageList from './components/MessageList';

const App = () => {
  return (
    <div className="bg-gray-900 min-h-screen text-white">
      <MessageList />
      <MessageForm />
    </div>
  );
};

export default App;
