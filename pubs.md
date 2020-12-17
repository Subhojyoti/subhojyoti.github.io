---
layout: page
title: Publications
---


# Peer-reviewed Conferences and Journals


1. *"A Unified Approach to Translate Classical Bandit Algorithms to the Structured Bandit Setting"*, Samarth Gupta, Shreyas Chaudhari, Subhojyoti Mukherjee, Gauri Joshi, Osman Yağan, **Special Issue on Estimation and Inference paper in the IEEE Journal on Selected Areas in Information Theory**.

   * ### Abstract ###
   
      <details>
        <summary>
          Show details
        </summary>
          <p>We consider a finite-armed structured bandit problem in which mean rewards of different arms are known functions of a common hidden parameter θ∗. Since we do not place any restrictions of these functions, the problem setting subsumes several previously studied frameworks that assume linear or invertible reward functions. We propose a novel approach to gradually estimate the hidden θ∗ and use the estimate together with the mean reward functions to substantially reduce exploration of sub-optimal arms. This approach enables us to fundamentally generalize any classic bandit algorithm including UCB and Thompson Sampling to the structured bandit setting. We prove via regret analysis that our proposed UCB-C and TS-C algorithms (structured bandit versions of UCB and Thompson Sampling, respectively) pull only a subset of the sub-optimal arms O(logT) times while the other sub-optimal arms (referred to as non-competitive arms) are pulled O(1) times. As a result, in cases where all sub-optimal arms are non-competitive, which can happen in many practical scenarios, the proposed algorithms achieve bounded regret. We also conduct simulations on the Movielens recommendations dataset to demonstrate the improvement of the proposed algorithms over existing structured bandit algorithms.</p>
      </details>
   
   * ### Full Paper ###
   
      [**PDF**](https://ieeexplore.ieee.org/document/9276444)

2. *"Efficient UCBV: An Almost Optimal Algorithm using Variance Estimates"*, Subhojyoti Mukherjee, K.P. Naveen, Nandan Sudarsanam, and Balaraman Ravindran, **Proceedings of the Thirty-Second Association for the Advancement of Artificial Intelligence (AAAI-18)**.

   * ### Abstract ###
   
      <details>
        <summary>
          Show details
        </summary>
          <p>We propose a novel variant of the UCB algorithm (referred to as Efficient-UCB-Variance (EUCBV)) for minimizing cumulative regret in the stochastic multi-armed bandit (MAB) setting. EUCBV incorporates the arm elimination strategy proposed in UCB-Improved while taking into account the variance estimates to compute the arms' confidence bounds, similar to UCBV. Through a theoretical analysis, we establish the *gap-dependent* regret bound of EUCBV after T trials and this bound is an improvement over that of existing state-of-the-art UCB algorithms (such as UCB1, UCB-Improved, UCBV,  MOSS). Further, EUCBV incurs an order optimal *gap-independent* regret bound of which is an improvement over that of UCB1, UCBV, and UCB-Improved, while being comparable with that of MOSS and OCUCB. Through an extensive numerical study, we show that EUCBV significantly outperforms the popular UCB variants (like MOSS, OCUCB, etc.) as well as Thompson sampling and Bayes-UCB algorithms.</p>
      </details>
   
   * ### Full Paper ###
   
      [**PDF**](https://www.aaai.org/ocs/index.php/AAAI/AAAI18/paper/view/16111)

3. *"Thresholding Bandits with Augmented UCB"*, Subhojyoti Mukherjee, K. P. Naveen, Nandan Sudarsanam, Balaraman Ravindran, **Proceedings of the Twenty-Sixth International Joint Conference on Artificial Intelligence (IJCAI-17)**.
   
   * ### Abstract ###
   
      <details>
      <summary>
          Show details
      </summary>
        <p>In this paper we propose the Augmented-UCB (AugUCB) algorithm for a fixed-budget version of the thresholding bandit problem  (TBP), where the objective is to identify a set of arms whose quality is above a threshold. A key feature of AugUCB is that it uses both  mean and variance estimates to eliminate arms that have been sufficiently explored; to the best of our knowledge, this is the first algorithm to employ such an approach for the considered TBP. Theoretically, we obtain an upper bound on the loss (probability of misclassification) incurred by AugUCB. Although UCBEV in literature provides a better guarantee, it is important to emphasize that UCBEV has access to problem complexity (whose computation requires arms’ mean and variances), and hence is not realistic in practice; this is in contrast to AugUCB whose implementation does not require any such complexity inputs. We conduct extensive simulation experiments to validate the performance of AugUCB. Through our simulation work, we establish that AugUCB, owing to its utilization of variance estimates, performs significantly better than the state-of-the-art APT, CSAR and other nonvariance-based algorithms.</p>
        </details>
   * ### Full Paper ###  
   
      [**PDF**](https://www.ijcai.org/proceedings/2017/0350.pdf)


# Peer-reviewed Workshops

1. *"Generalized Chernoff Sampling for Active Learning and Structured Bandit Algorithms"*, Subhojyoti Mukherjee, Ardhendu Tripathy, Robert Nowak, **An initial version was accepted at the Theoretical Foundations of Reinforcement Learning Workshop in the 37 th International Conference on Machine Learning, 2020 (ICML-20)**.
   
   * ### Abstract ###
   
      <details>
      <summary>
          Show details
      </summary>
        <p>Active learning and structured stochastic bandit problems are intimately related to the classical problem of sequential experimental design. This paper studies active learning and best-arm identification in structured bandit settings from the viewpoint of active sequential hypothesis testing, a framework initiated by Chernoff (1959). We first characterize the sample complexity of Chernoff's original procedure by uncovering terms that reduce in significance as the allowed error probability δ→0, but are nevertheless relevant at any fixed value of δ>0. While initially proposed for testing among finitely many hypotheses, we obtain the analogue of Chernoff sampling for the case when the hypotheses belong to a compact space. This makes it applicable to active learning and structured bandit problems, where the unknown parameter specifying the arm means is often assumed to be an element of Euclidean space. Empirically, we demonstrate the potential of our proposed approach for active learning of neural network models and in the linear bandit setting, where we observe that our general-purpose approach compares favorably to state-of-the-art methods.</p>
        </details>
   * ### Full Paper ###  
   
      [**PDF**](https://arxiv.org/abs/2012.08073)

2. *"Distribution-dependent and Time-uniform Bounds for Piecewise i.i.d Bandits"*, Subhojyoti Mukherjee, Odalric-Ambrym Maillard, **An initial version was accepted at the Reinforcement Learning for Real Life (RL4RealLife) Workshop in the 36 th International Conference on Machine Learning, Long Beach, California, USA, 2019 (ICML-19)**.
   
   * ### Abstract ###
   
      <details>
      <summary>
          Show details
      </summary>
        <p>We consider the setup of stochastic multi-armed bandits in the case when reward distributions are piecewise i.i.d. and bounded with unknown changepoints. We focus on the case when changes happen simultaneously on all arms, and in stark contrast with the existing literature, we target gap-dependent (as opposed to only gap-independent) regret bounds involving the magnitude of changes and optimality-gaps. Diverging from previous works, we assume the more realistic scenario that there can be undetectable changepoint gaps and under a different set of assumptions, we show that as long as the compounded delayed detection for each changepoint is bounded there is no need for forced exploration to actively detect changepoints. We introduce two adaptations of UCB-strategies that employ scan-statistics in order to actively detect the changepoints, without knowing in advance the changepoints and also the mean before and after any change. Our first method UCBLCPD does not know the number of changepoints G or time horizon T and achieves the first time-uniform concentration bound for this setting using the Laplace method of integration. The second strategy ImpCPD makes use of the knowledge of T to achieve the order optimal regret bound thereby closing an important gap with respect to the lower bound in a specific challenging setting. Our theoretical findings are supported by numerical experiments on synthetic and real-life datasets.</p>
        </details>
   * ### Full Paper ###  
   
      [**PDF**](https://arxiv.org/abs/1905.13159)

# Technical Reports

1. *Latent Ranked Bandits*, Subhojyoti Mukherjee, Branislav Kveton, Anup Rao

   * ### Abstract ###
   
      <details>
        <summary>
          Show details
        </summary>
          <p> We study the problem of learning personalized ranked lists of diverse items for multiple users, from sequential observations of user preferences. The user-item preference matrix is non-negative and low-rank. Existing methods for solving similar problems are based on reconstructing the preference matrix from its noisy observations using matrix factorization techniques, and typically require strong assumptions on the reconstructed matrix. We depart from this standard approach and consider a family of low-rank matrices, where the set of most preferred items of all users is small and can be learned efficiently. Then we learn to present this set to each user in a personalized manner, in the order of the descending preferences of the user. We propose a computationally efficient algorithm that implements this procedure, and prove a sublinear bound on its n-step regret. We evaluate the algorithm empirically on several synthetic and real-world datasets. In all experiments, we outperform existing state-of-the-art algorithms. </p>
       </details>
    
    
       [Download Report](/pdf/paper.pdf)
       


