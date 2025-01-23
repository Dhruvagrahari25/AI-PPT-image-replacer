import React, { useState } from 'react';
import axios from 'axios';

const MessageForm = ({ onAddMessage }) => {
  const [text, setText] = useState('');
  const [username, setUsername] = useState('');

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!text.trim() || !username.trim()){
        window.alert("UserName or message missing...!");
    }

    const messageData = {
      username,
      text,
      timestamp: new Date().toISOString(),
    };

    try {
      const response = await axios.post('http://localhost:5000/messages', messageData);
      setText('');
      onAddMessage(response.data); // Assuming the API returns the created message
    } catch (error) {
      console.error('Error sending message:', error);
    }
  };

  return (
    <form onSubmit={handleSubmit} className="fixed bottom-0 left-0 w-full bg-gray-900 border-t border-none shadow-lg">
      <div className="max-w-2xl mx-auto flex items-center space-x-4 p-3">
        <input
          type="text"
          placeholder="Your name"
          value={username}
          onChange={(e) => setUsername(e.target.value)}
          className="p-3 border border-gray-700 rounded-md bg-gray-800 text-white w-40 focus:outline-none"
        />
        <textarea
          className="flex-1 p-3 border border-gray-700 rounded-md bg-gray-800 text-white focus:outline-none"
          rows="1"
          placeholder="Type your message..."
          value={text}
          onChange={(e) => setText(e.target.value)}
        />
        <button type="submit" className="bg-green-600 text-white px-6 py-3 rounded-md hover:bg-green-700 transition duration-300">
          Post
        </button>
      </div>
    </form>
  );
};

export default MessageForm;
