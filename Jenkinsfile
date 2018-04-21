#!/usr/bin/groovy

def buildProject(String distro, String buildType) {
  node(distro) {
    deleteDir()
    checkout scm
    sh """
    distribution=\$(lsb_release -is)
    distributionVersion=\$(lsb_release -rs)

    if [ "\$distribution" = "Ubuntu" ]; then
        if [ "\$distributionVersion" = "16.04" ]; then
            sudo apt-get install -y \
                cmake \
                build-essential \
                g++ \
                libcurl4-openssl-dev \
                libasan1
        fi
    fi

    if [ "\$distribution" = "CentOS" ]; then
        sudo yum install -y \
            cmake \
            gcc-c++ \
            libcurl-devel \
            libasan
    fi
    """

    withCredentials([usernamePassword(credentialsId: 'c0ede5c0-82be-4ff0-b618-39aeddcdfaba', passwordVariable: 'password', usernameVariable: 'username')]) {
      dir('build') {
        sh """
        cmake -DCMAKE_BUILD_TYPE=${buildType} -DENABLE_SANITIZERS=ON -DDISABLE_TLS_CERT_VERIFICATION=ON ..
        make -j2

        EWS_TEST_URI=https://duckburg.otris.de/ews/Exchange.asmx \
        EWS_TEST_DOMAIN=OTRIS \
        EWS_TEST_USERNAME=${username}\
        EWS_TEST_PASSWORD=${password}\
        ./tests --assets=${WORKSPACE}/tests/assets --gtest_output=xml:result.xml||true
        """

        step([$class: 'XUnitBuilder',
              testTimeMargin: '3000',
              thresholdMode: 1,
              thresholds: [[$class: 'FailedThreshold', failureNewThreshold: '1', failureThreshold: '1', unstableNewThreshold: '1', unstableThreshold: '1'],
              [$class: 'SkippedThreshold', failureNewThreshold: '', failureThreshold: '', unstableNewThreshold: '', unstableThreshold: '']],
              tools: [[$class: 'GoogleTestType', deleteOutputFiles: true, failIfNotNew: true, pattern: 'result.xml', skipNoTestFiles: false, stopProcessingIfError: true]]])
      }
    }
  }
}

pipeline {
  agent { label 'Linux-SCM' }
  options {
    buildDiscarder(logRotator(numToKeepStr: '5'))
  }
  parameters {
    booleanParam(name: 'BUILD_UBUNTU1604', defaultValue: true, description: 'Build for Ubuntu 16.04')
    booleanParam(name: 'BUILD_CENTOS7', defaultValue: true, description: 'Build for Centos 7')
  }
  triggers {
    pollSCM('@hourly')
  }
  stages {
    stage ('Build & Test') {
      parallel {
        stage('Build Ubuntu 16.04') {
          when {
            expression {
              return params.BUILD_UBUNTU1604
            }
          }
          steps {
            buildProject('ubuntu1604-64-cloud', 'RelWithDebInfo')
          }
        }
        stage('Build on Centos 7') {
          when {
            expression {
              return params.BUILD_CENTOS7
            }
          }
          steps {
            buildProject('centos7-64-cloud', 'RelWithDebInfo')
          }
        }
      }
    }
  }
}
