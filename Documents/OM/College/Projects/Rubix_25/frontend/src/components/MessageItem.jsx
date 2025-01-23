import React from 'react';

const MessageItem = ({ text }) => {
  return (
    <div className="p-4 bg-gray-800 border border-gray-700 rounded-lg shadow-md">
      <p className="text-gray-300">{text}</p>
      <div className="mt-2 text-right text-sm text-gray-500">- Anonymous</div>
    </div>
  );
};

export default MessageItem;
