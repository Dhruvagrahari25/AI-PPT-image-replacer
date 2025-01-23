import React from 'react';

const Footer = () => {
  return (
    <footer className="bg-gray-800 text-gray-200 py-8">
      <div className="max-w-6xl mx-auto text-center">
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-6">
          {/* Column 1 */}
          <div>
            <h3 className="text-xl font-bold text-green-00 mb-3">About Us</h3>
            <p className="text-sm">
              Our forum is a community-driven platform for open discussions, sharing knowledge, and connecting with like-minded individuals.
            </p>
          </div>
          {/* Column 2 */}
          <div>
            <h3 className="text-xl font-bold text-green-00 mb-3">Quick Links</h3>
            <ul className="space-y-2">
              <li>
                <a href="/forum" className="hover:text-green-00 transition">
                  Forum
                </a>
              </li>
              <li>
                <a href="#features" className="hover:text-green-00 transition">
                  Features
                </a>
              </li>
              <li>
                <a href="#testimonials" className="hover:text-green-00 transition">
                  Testimonials
                </a>
              </li>
              <li>
                <a href="#contact" className="hover:text-green-00 transition">
                  Contact Us
                </a>
              </li>
            </ul>
          </div>
          {/* Column 3 */}
          <div>
            <h3 className="text-xl font-bold text-green-00 mb-3">Contact</h3>
            <p className="text-sm">
              Email: <a href="mailto:support@communityforum.com" className="hover:text-green-00 transition">support@communityforum.com</a>
            </p>
            <p className="text-sm">Phone: +1-234-567-8901</p>
            <p className="text-sm">Address: 123 Forum Street, Knowledge City</p>
          </div>
        </div>

        <p className="text-sm text-gray-400 mt-6">
          &copy; {new Date().getFullYear()} Community Forum. All rights reserved.
        </p>
      </div>
    </footer>
  );
};

export default Footer;
