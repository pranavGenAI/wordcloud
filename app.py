import streamlit as st
import pandas as pd
from wordcloud import WordCloud, STOPWORDS
import matplotlib.pyplot as plt
import io

st.set_page_config(page_title="Excel WordCloud Generator", page_icon="📊", layout="wide")

particle_html = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Particles</title>
    <style>
        body {
            margin: 0;
            overflow: hidden;
        }
        #tsparticles {
            position: absolute;
            width: 100%;
            height: 100%;
            background: #000;
        }
    </style>
</head>
<body>
    <div id="tsparticles"></div>
    <script src="https://cdn.jsdelivr.net/npm/tsparticles@2.1.0/tsparticles.bundle.min.js"></script>
    <script>
        tsParticles.load("tsparticles", {
            particles: {
                number: {
                    value: 80,
                    density: {
                        enable: true,
                        value_area: 800
                    }
                },
                color: {
                    value: "#f29fff"
                },
                shape: {
                    type: "circle",
                    stroke: {
                        width: 0,
                        color: "#000000"
                    },
                    polygon: {
                        nb_sides: 5
                    }
                },
                opacity: {
                    value: 0.5,
                    random: false,
                    anim: {
                        enable: false,
                        speed: 1,
                        opacity_min: 0.1,
                        sync: false
                    }
                },
                size: {
                    value: 3,
                    random: true,
                    anim: {
                        enable: false,
                        speed: 40,
                        size_min: 0.1,
                        sync: false
                    }
                },
                line_linked: {
                    enable: true,
                    distance: 150,
                    color: "#ffffff",
                    opacity: 0.4,
                    width: 1
                },
                move: {
                    enable: true,
                    speed: 1,
                    direction: "none",
                    random: false,
                    straight: false,
                    out_mode: "out",
                    bounce: false,
                    attract: {
                        enable: false,
                        rotateX: 600,
                        rotateY: 1200
                    }
                }
            },
            interactivity: {
                detect_on: "canvas",
                events: {
                    onhover: {
                        enable: true,
                        mode: "grab"
                    },
                    onclick: {
                        enable: true,
                        mode: "push"
                    },
                    resize: true
                },
                modes: {
                    grab: {
                        distance: 200,
                        line_linked: {
                            opacity: 1
                        }
                    },
                    bubble: {
                        distance: 400,
                        size: 40,
                        duration: 2,
                        opacity: 8,
                        speed: 3
                    },
                    repulse: {
                        distance: 200,
                        duration: 0.4
                    },
                    push: {
                        particles_nb: 4
                    },
                    remove: {
                        particles_nb: 2
                    }
                }
            },
            retina_detect: true
        });
    </script>
</body>
</html>
"""

# Embed the HTML code into the Streamlit app
st.components.v1.html(particle_html, height=1000)

# Add CSS to make the iframe fullscreen
st.markdown("""
<style>
    iframe {
        position: fixed;
        left: 0;
        right: 0;
        top: 0;
        bottom: 0;
        border: none;
        height: 100%;
        width: 100%;
    }
</style>
""", unsafe_allow_html=True)

# Add CSS to make the iframe fullscreen and hide the viewer badge
st.markdown("""
<style>
    iframe {
        position: fixed;
        left: 0;
        right: 0;
        top: 0;
        bottom: 0;
        border: none;
        height: 100%;
        width: 100%;
    }

    .viewerBadge_container__r5tak {
        display: none;
    }
</style>
""", unsafe_allow_html=True)

col1, col2, col3 = st.columns([1, 50, 1])

with col1:
    st.write(' ')

with col2:
    st.markdown("""
        <style>
            @keyframes gradientAnimation {
                0% { background-position: 0% 50%; }
                50% { background-position: 100% 50%; }
                100% { background-position: 0% 50%; }
            }

            .animated-gradient-text {
                font-family: "Graphik Semibold";
                font-size: 42px;
                background: linear-gradient(45deg, rgb(245, 58, 126) 30%, rgb(200, 1, 200) 55%, rgb(197, 45, 243) 20%);
                background-size: 300% 200%;
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
                animation: gradientAnimation 6s ease-in-out infinite;
                color: #FFF;
                transition: color 0.5s, text-shadow 0.5s;
            }

            @keyframes glow {
                0%, 18%, 20%, 50.1%, 60%, 65.1%, 80%, 90.1%, 92% {
                    color: #0e3742;
                    text-shadow: none;
                }
                18.1%, 20.1%, 30%, 50%, 60.1%, 65%, 80.1%, 90%, 92.1%, 100% {
                    color: #fff;
                    text-shadow: 0 0 10px rgb(197, 45, 243), 0 0 20px rgb(197, 45, 243);
                }
            }

            .animated-gradient-text:hover {
                animation: glow 5s linear infinite;
            }

            .glow-on-hover {
                transition: transform 0.5s, filter 0.3s;
            }

            .glow-on-hover:hover {
                transform: scale(1.15);
                filter: drop-shadow(0 0 10px rgba(197, 45, 243, 0.8));
            }
        </style>
        <p class="animated-gradient-text" align="center">
            Excel WordCloud Generator
        </p>
    """, unsafe_allow_html=True)

# Function to generate wordcloud
def generate_wordcloud(text, exclude_words, include_words=None):
    stopwords = set(STOPWORDS)  # Built-in stopwords
    stopwords.update(exclude_words)  # Add user-defined exclude words to stopwords

    if include_words:
        # Include only the words specified by the user
        word_list = [word for word in text.split() if word.lower() in include_words]
        text = " ".join(word_list)

    wordcloud = WordCloud(
        width=800,
        height=400,
        background_color='white',
        stopwords=stopwords,  # Remove common stopwords and user-defined words
        max_words=200,         # Limit to top 200 words
        colormap='viridis',    # Use a color map for better aesthetics
        contour_color='black',
        contour_width=2        # Add contour to the word cloud
    ).generate(text)
    
    return wordcloud

# Streamlit App Layout

# Title of the app
st.markdown("""
    <h3 style="font-size: 24px;">Upload an Excel file  &gt;  Select a column for word extraction  &gt;  Input words to include/exclude from the word cloud  &gt;  Generate Wordcloud</h3>
""", unsafe_allow_html=True)

# File upload
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Read the Excel file
    df = pd.read_excel(uploaded_file)

    # Display dataframe to user for reference
    st.subheader("Uploaded Excel Data")
    st.write(df.head())  # Show first 5 rows to the user

    # Ask user to input the column they want to use (default is column B, which is the second column)
    column_name = st.selectbox("Select the column to use for word extraction", df.columns)

    # Ask the user for words to exclude from the word cloud (comma-separated)
    exclude_words_input = st.text_input("Enter words to exclude (comma-separated)")

    if exclude_words_input:
        exclude_words = set(exclude_words_input.split(","))
    else:
        exclude_words = set()

    # Ask the user for words they want to include in the word cloud (comma-separated)
    include_words_input = st.text_input("Enter words to include (comma-separated, optional)")

    include_words = set(include_words_input.split(",")) if include_words_input else set()

    # Alternatively, allow the user to multi-select words from the column
    word_list = set(" ".join(df[column_name].dropna().astype(str)).split())
    include_words_select = st.multiselect("Select words to include from the word list", options=list(word_list))

    if include_words_select:
        include_words.update(set(include_words_select))

    # Generate word cloud when the user is ready
    if st.button("Generate WordCloud"):
        # Extract text from the selected column
        text = " ".join(df[column_name].dropna().astype(str))

        # Generate the wordcloud with exclusions and inclusions
        wordcloud = generate_wordcloud(text, exclude_words, include_words)

        # Display the wordcloud
        st.subheader(f"Word Cloud for Column '{column_name}'")
        plt.figure(figsize=(8, 8), facecolor=None)
        plt.imshow(wordcloud, interpolation="bilinear")
        plt.axis("off")
        st.pyplot(plt)
