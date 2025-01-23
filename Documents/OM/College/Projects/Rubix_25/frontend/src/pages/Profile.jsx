import React, { useState, useEffect } from "react";
import axios from "axios";

const ProfilePage = () => {
  const [userData, setUserData] = useState(null);

  // Fetch user data from the backend
  useEffect(() => {
    axios
      .get("/api/user/profile") // Adjust API endpoint as needed
      .then((response) => {
        setUserData(response.data);
      })
      .catch((error) => console.error("Error fetching user data:", error));
  }, []);

  if (!userData) {
    return (
      <div className="flex items-center justify-center min-h-screen text-gray-700">
        Loading...
      </div>
    );
  }

  const { name, profileImage, purchaseHistory, carbonFootprintSaved, email } =
    userData;

  return (
    <div className="bg-gradient-to-br from-gray-800 via-gray-900 to-black text-white min-h-screen py-8 px-6">
      <div className="max-w-4xl mx-auto bg-gray-800 rounded-lg shadow-lg p-6">
        {/* Profile Header */}
        <div className="flex items-center mb-8">
          <img
            src={profileImage}
            alt={`${name}'s profile`}
            className="w-24 h-24 rounded-full border-4 border-green-500 shadow-md"
          />
          <div className="ml-6">
            <h1 className="text-3xl font-bold">{name}</h1>
            <p className="text-gray-400">{email}</p>
          </div>
        </div>

        {/* Metrics Section */}
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-6 mb-8">
          <div className="bg-green-700 text-center p-6 rounded-lg shadow-md">
            <h2 className="text-2xl font-semibold">Carbon Footprint Saved</h2>
            <p className="text-4xl font-bold mt-2">
              {carbonFootprintSaved} kg CO₂
            </p>
          </div>
          <div className="bg-blue-700 text-center p-6 rounded-lg shadow-md">
            <h2 className="text-2xl font-semibold">Purchases</h2>
            <p className="text-4xl font-bold mt-2">{purchaseHistory.length}</p>
          </div>
        </div>

        {/* Purchase History Table */}
        <h2 className="text-2xl font-bold mb-4">Purchase History</h2>
        <div className="overflow-x-auto">
          <table className="w-full table-auto bg-gray-700 rounded-lg">
            <thead>
              <tr className="text-left bg-gray-900 text-green-500">
                <th className="px-4 py-2">Date</th>
                <th className="px-4 py-2">Product</th>
                <th className="px-4 py-2">Carbon Saved</th>
              </tr>
            </thead>
            <tbody>
              {purchaseHistory.map((purchase, index) => (
                <tr
                  key={index}
                  className="hover:bg-gray-800 border-b border-gray-600"
                >
                  <td className="px-4 py-2">{purchase.date}</td>
                  <td className="px-4 py-2">{purchase.product}</td>
                  <td className="px-4 py-2">{purchase.carbonSaved} kg CO₂</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

export default ProfilePage;
