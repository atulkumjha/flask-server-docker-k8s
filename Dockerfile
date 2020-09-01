FROM  continuumio/anaconda3
COPY . /app
WORKDIR /app
RUN pip install -r requirements.txt
ENV FLASK_APP=run_sample_app
ENV PATH /opt/conda/bin:$PATH
COPY boot.sh ./
RUN chmod +x boot.sh
CMD ./boot.sh
