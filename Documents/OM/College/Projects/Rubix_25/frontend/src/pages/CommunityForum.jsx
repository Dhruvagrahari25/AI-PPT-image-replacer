import React, { useState, useEffect } from 'react';
import MessageForm from '../components/MessageForm';
import MessageList from '../components/MessageList';
import axios from 'axios';

const CommunityForum = () => {
  const [messages, setMessages] = useState([]);

  useEffect(() => {
    axios.get('http://localhost:5000/messages')
      .then(response => setMessages(response.data))
      .catch(error => console.error('Error fetching messages:', error));
  }, []);

  const addMessage = (newMessage) => {
    setMessages([...messages, newMessage]);
  };

  return (
    <div className='m-5'>
      {/* <h1>Community Forum</h1> */}
      <MessageList messages={messages} />
      <MessageForm onAddMessage={addMessage} />
    </div>
  );
};

export default CommunityForum;