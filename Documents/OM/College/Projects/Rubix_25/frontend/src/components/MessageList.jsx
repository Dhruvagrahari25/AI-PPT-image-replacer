import React, { useEffect, useState } from 'react';
import axios from 'axios';

const MessageList = () => {
  const [messages, setMessages] = useState([]);

  useEffect(() => {
    const fetchMessages = async () => {
      try {
        const response = await axios.get('http://localhost:5000/messages');
        setMessages(response.data);
      } catch (error) {
        console.error('Error fetching messages:', error);
      }
    };

    fetchMessages();
    const interval = setInterval(fetchMessages, 5000); // Poll every 5 seconds

    return () => clearInterval(interval); // Cleanup on unmount
  }, []);

  return (
    <div className="max-w-2xl mx-auto mt-8 mb-20">
      <h2 className="text-3xl font-bold text-green-400 mb-4 text-center">Community Messages</h2>
      <div className="space-y-4">
        {messages.length > 0 ? (
          messages.map((msg, index) => (
            <div key={index} className="p-4 bg-gray-800 border border-gray-700 rounded-lg shadow-md">
              <p className="text-gray-300">{msg.text}</p>
              <div className="mt-2 text-sm text-gray-500">
                - {msg.username} at {new Date(msg.timestamp).toLocaleTimeString()}
              </div>
            </div>
          ))
        ) : (
          <p className="text-center text-gray-400">No messages yet. Be the first to post!</p>
        )}
      </div>
    </div>
  );
};

export default MessageList;
