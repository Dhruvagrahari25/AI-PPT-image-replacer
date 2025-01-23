import { useState } from "react";
import logo from '../assets/Images/logo.png'

export default function Navbar() {
  const [isOpen, setIsOpen] = useState(false);

  return (
    <nav className="bg-transparent text-white">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="flex justify-between items-center h-18 my-3">
          {/* Logo */}
          <div className="flex items-center">
            <img src={logo} alt="Logo" className="h-12" />
          </div>
          {/* Hamburger Menu for Mobile */}
          <div className="sm:hidden">
            <button
              className="p-2 rounded-md focus:outline-none"
              onClick={() => setIsOpen(!isOpen)}
            >
              {isOpen ? "✖" : "☰"}
            </button>
          </div>
          {/* Navigation Links */}
          <div
            className={`sm:flex items-center space-x-4 ${
              isOpen ? "block" : "hidden"
            }`}
          >
            <a href="#" className="block px-3 py-2 hover:text-gray-300">
              Home
            </a>
            <a href="#" className="block px-3 py-2 hover:text-gray-300">
              About
            </a>
            <a href="#" className="block px-3 py-2 hover:text-gray-300">
              Contact
            </a>
          </div>
        </div>
      </div>
    </nav>
  );
}
