import React from 'react';

const TestimonialsSection = () => {
  const testimonials = [
    {
      name: 'Alex Periera',
      feedback: 'This forum is a fantastic place to connect with others and share knowledge. Highly recommend!',
      avatar: 'https://www.gravatar.com/avatar/2c7d99fe281ecd3bcd65ab915bac6dd5?s=250', // Replace with an actual avatar URL
    },
    {
      name: 'Connor McGregor',
      feedback: 'I’ve learned so much here. The community is friendly, and the discussions are very insightful.',
      avatar: 'https://mdbcdn.b-cdn.net/img/new/avatars/2.webp', // Replace with an actual avatar URL
    },
    {
      name: 'Valentina Shevchenko',
      feedback: 'A great platform for meaningful conversations. It has really helped me grow!',
      avatar: 'https://mdbcdn.b-cdn.net/img/Photos/Avatars/img%20(1).webp', // Replace with an actual avatar URL
    },
  ];

  return (
    <div className="bg-gray-900 text-white py-28 px-6">
      <div className="max-w-5xl mx-auto text-center">
        <h2 className="text-5xl font-bold text-green-00 mb-8">What Our Members Say</h2>
        <p className="text-lg text-gray-200 mb-12">
          Hear from some of the amazing members of our community!
        </p>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
          {testimonials.map((testimonial, index) => (
            <div
              key={index}
              className="bg-gray-800 p-6 rounded-lg shadow-lg text-center"
            >
              <img
                src={testimonial.avatar}
                alt={testimonial.name}
                className="w-20 h-20 rounded-full mx-auto mb-4"
              />
              <h3 className="text-xl font-semibold mb-2">{testimonial.name}</h3>
              <p className="text-gray-300 italic">"{testimonial.feedback}"</p>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

export default TestimonialsSection;
