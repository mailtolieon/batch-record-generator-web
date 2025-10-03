#!/bin/bash

echo "🚀 Setting up Batch Record Generator Web Application"

# Create virtual environment
echo "📦 Creating virtual environment..."
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
echo "📥 Installing dependencies..."
pip install -r requirements.txt

# Create necessary directories
echo "📁 Creating directories..."
mkdir -p data generated temp

echo "✅ Setup complete!"
echo ""
echo "🎯 To run the application locally:"
echo "   streamlit run app.py"
echo ""
echo "🎯 To deploy to Streamlit Cloud:"
echo "   1. Push this code to GitHub"
echo "   2. Go to https://share.streamlit.io"
echo "   3. Connect your repository"
echo "   4. Set main file to app.py"
echo ""
echo "🎯 To run with Docker:"
echo "   docker build -t batch-record-generator ."
echo "   docker run -p 8501:8501 batch-record-generator"