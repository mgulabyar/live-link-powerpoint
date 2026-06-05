// Image ko compress aur resize karne ke liye helper
export const processAndCompressImage = (base64Str: string): Promise<string> => {
  return new Promise((resolve) => {
    const img = new Image();
    img.src = base64Str.startsWith("data:") ? base64Str : `data:image/png;base64,${base64Str}`;
    
    img.onload = () => {
      const canvas = document.createElement("canvas");
      let width = img.width;
      let height = img.height;

      // Agar image 1200px se bari hai to scale down karein (PowerPoint safety)
      const maxSide = 1200;
      if (width > maxSide || height > maxSide) {
        if (width > height) {
          height = (maxSide / width) * height;
          width = maxSide;
        } else {
          width = (maxSide / height) * width;
          height = maxSide;
        }
      }

      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext("2d");
      if (ctx) {
        ctx.fillStyle = "white"; // Background white rakhein
        ctx.fillRect(0, 0, width, height);
        ctx.drawImage(img, 0, 0, width, height);
      }

      // PNG ko JPEG mein convert karein (JPEG ka size 10 guna chota hota hai)
      const compressedBase64 = canvas.toDataURL("image/jpeg", 0.8);
      // Sirf raw part return karein
      resolve(compressedBase64.split(",")[1]);
    };
    
    img.onerror = () => resolve("");
  });
};