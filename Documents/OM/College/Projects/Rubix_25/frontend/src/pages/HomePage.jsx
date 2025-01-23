import React from 'react';
import { Link } from 'react-router-dom';
import HeroSection from '../components/HeroSection';
import FeaturesSection from '../components/FeatureSection';
import TestimonialsSection from '../components/Testimonials';
import Footer from '../components/Footer';
import Navbar from '../components/Navbar';

const HomePage = () => {
  return (
    // <div className="bg-gray-900 text-white min-h-screen flex flex-col items-center justify-center">
    //   <header className="text-center">
    //     <h1 className="text-5xl font-bold text-green-500 mb-6">Welcome to the Community Forum</h1>
    //     <p className="text-lg text-gray-300 max-w-lg mx-auto">
    //       Connect, share, and grow with a vibrant community. Post your thoughts, read messages, and be part of the conversation!
    //     </p>
    //   </header>

    //   <div className="mt-10 flex space-x-6">
    //     <Link
    //       to="/forum"
    //       className="bg-green-500 text-white px-8 py-4 rounded-lg text-lg font-semibold shadow-lg hover:bg-green-600 transition"
    //     >
    //       Go to Forum
    //     </Link>
    //     <button
    //       className="bg-gray-800 px-8 py-4 rounded-lg text-lg font-semibold shadow-lg hover:bg-gray-700 transition"
    //       onClick={() => alert('More features coming soon!')}
    //     >
    //       Learn More
    //     </button>
    //   </div>

    //   <footer className="absolute bottom-4 text-gray-500 text-sm">
    //     &copy; {new Date().getFullYear()} Community Forum. All rights reserved.
    //   </footer>
    // </div>
    <>
        <Navbar />
        <HeroSection />
        <FeaturesSection />
        <TestimonialsSection />
        <Footer />
    </>
  );
};

export default HomePage;
