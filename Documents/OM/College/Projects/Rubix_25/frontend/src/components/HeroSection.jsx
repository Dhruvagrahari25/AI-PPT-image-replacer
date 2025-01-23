import React from "react";
import { Link } from "react-router-dom";
import herosectionImg from "../assets/Images/herosectionImg.jpeg";

const HeroSection = () => {
  return (
    <div className="bg-gray-900 text-white min-h-screen flex flex-col lg:flex-row items-center justify-center text-center lg:text-left p-6 space-y-8 lg:space-y-0 lg:space-x-10">
      {/* Text Section */}
      <div className="flex flex-col items-center lg:items-start max-w-2xl">
        <h1 className="text-4xl sm:text-5xl font-extrabold mb-4">
          Empower Your Voice for{" "}
          <span className="text-green-600">Mother Nature</span>
        </h1>
        <p className="text-lg sm:text-xl text-gray-200 mb-8">
          A place where ideas are shared, connections are built, and inspiration
          is sparked. Your thoughts matter, and together, we can make a
          difference.
        </p>
        <div className="flex flex-col sm:flex-row space-y-4 sm:space-y-0 sm:space-x-6">
          <Link
            to="/forum"
            className="bg-green-800 text-white px-8 py-4 rounded-lg font-semibold shadow-lg hover:bg-green-400 transition duration-300"
          >
            Explore Eco-friendly Choices
          </Link>
          <a
            href="http://localhost:5173/profile"
            className="bg-white text-gray-900 px-8 py-4 rounded-lg font-semibold shadow-lg hover:bg-gray-200 transition duration-300"
          >
            Discover More
          </a>
        </div>
      </div>

      {/* Image Section */}
      <div className="flex justify-center">
        <img
          src={herosectionImg}
          alt="Community Inspiration"
          className="rounded-lg shadow-lg w-80 h-auto lg:w-[500px]"
        />
      </div>
    </div>
  );
};

export default HeroSection;
