import React from 'react';
import { FaUsers, FaComments, FaRocket } from 'react-icons/fa'; // Icons for features

const FeaturesSection = () => {
  return (
    <div
      id="features"
      className="bg-gradient-to-br from-gray-900 via-gray-800 to-black text-white py-16 px-6"
    >
      <div className="max-w-5xl mx-auto text-center">
        {/* Title and Description */}
        <h2 className="text-4xl sm:text-5xl font-bold text-green-500 mb-8">
          Why Join Us?
        </h2>
        <p className="text-md sm:text-lg text-gray-200 mb-12">
          Discover what makes our community forum the perfect place to share
          your thoughts, connect with like-minded individuals, and grow
          together.
        </p>

        {/* Features Grid */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-8">
          {/* Feature 1 */}
          <div className="bg-gray-800 p-6 rounded-lg shadow-lg text-center">
            <FaUsers className="text-green-500 text-5xl mb-4 mx-auto" />
            <h3 className="text-xl sm:text-2xl font-semibold mb-2">
              Connect with People
            </h3>
            <p className="text-gray-300 text-sm sm:text-base">
              Meet a diverse community of individuals eager to share and learn
              together.
            </p>
          </div>

          {/* Feature 2 */}
          <div className="bg-gray-800 p-6 rounded-lg shadow-lg text-center">
            <FaComments className="text-green-500 text-5xl mb-4 mx-auto" />
            <h3 className="text-xl sm:text-2xl font-semibold mb-2">
              Engaging Discussions
            </h3>
            <p className="text-gray-300 text-sm sm:text-base">
              Dive into meaningful conversations and exchange ideas in real-time.
            </p>
          </div>

          {/* Feature 3 */}
          <div className="bg-gray-800 p-6 rounded-lg shadow-lg text-center">
            <FaRocket className="text-green-500 text-5xl mb-4 mx-auto" />
            <h3 className="text-xl sm:text-2xl font-semibold mb-2">
              Grow Together
            </h3>
            <p className="text-gray-300 text-sm sm:text-base">
              Share knowledge, gain insights, and grow both personally and
              professionally.
            </p>
          </div>
        </div>
      </div>
    </div>
  );
};

export default FeaturesSection;
