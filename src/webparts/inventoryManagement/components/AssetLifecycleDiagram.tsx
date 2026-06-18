import * as React from 'react';
import { Icon } from '@fluentui/react';

export interface IAssetLifecycleDiagramProps {
  isDarkTheme?: boolean;
}

export const AssetLifecycleDiagram: React.FC<IAssetLifecycleDiagramProps> = ({ isDarkTheme }) => {
  const [hoveredNode, setHoveredNode] = React.useState<number | null>(null);

  const nodes = [
    {
      id: 1,
      label: 'Request for an IT asset',
      icon: 'DocumentApproval',
      color: '#60a5fa', // Light blue
      // Top: (140, 50)
      left: 140,
      top: 50,
    },
    {
      id: 2,
      label: 'Acquisition of an asset',
      icon: 'PaymentCard',
      color: '#34d399', // Emerald/green
      // Top-right: (226, 112)
      left: 226,
      top: 112,
    },
    {
      id: 3,
      label: 'Deployment and assignment',
      icon: 'Settings',
      color: '#fb7185', // Rose/pink
      // Bottom-right: (193, 213)
      left: 193,
      top: 213,
    },
    {
      id: 4,
      label: 'Ongoing support & maintenance',
      icon: 'Headset',
      color: '#fbbf24', // Amber/yellow
      // Bottom-left: (87, 213)
      left: 87,
      top: 213,
    },
    {
      id: 5,
      label: 'End of life & disposition',
      icon: 'Recycle',
      color: '#a78bfa', // Purple
      // Top-left: (54, 112)
      left: 54,
      top: 112,
    },
  ];

  return (
    <div style={{
      position: 'relative',
      width: '280px',
      height: '280px',
      userSelect: 'none',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
    }}>
      {/* SVG Background for Connecting Ring */}
      <svg
        width="280"
        height="280"
        viewBox="0 0 280 280"
        style={{
          position: 'absolute',
          top: 0,
          left: 0,
          pointerEvents: 'none',
          zIndex: 1,
        }}
      >
        <defs>
          <linearGradient id="ringGradient" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" stopColor="#ffffff" stopOpacity={0.1} />
            <stop offset="50%" stopColor="#ffffff" stopOpacity={0.3} />
            <stop offset="100%" stopColor="#ffffff" stopOpacity={0.1} />
          </linearGradient>
          <style>{`
            @keyframes rotateRing {
              from { transform: rotate(0deg); }
              to { transform: rotate(360deg); }
            }
            .animated-ring {
              transform-origin: 140px 140px;
              animation: rotateRing 50s linear infinite;
            }
          `}</style>
        </defs>

        {/* Outer connecting circle (Dashed & Animated) */}
        <circle
          className="animated-ring"
          cx="140"
          cy="140"
          r="90"
          fill="none"
          stroke="url(#ringGradient)"
          strokeWidth="1.5"
          strokeDasharray="5, 6"
        />

        {/* Highlight ring on hover */}
        {hoveredNode !== null && (
          <circle
            cx="140"
            cy="140"
            r="90"
            fill="none"
            stroke="rgba(255, 255, 255, 0.35)"
            strokeWidth="1.5"
            strokeDasharray="8, 8"
            style={{
              transition: 'all 0.5s ease',
            }}
          />
        )}
      </svg>

      {/* Central Circle Badge (Glassmorphic) */}
      <div style={{
        position: 'absolute',
        width: '90px',
        height: '90px',
        borderRadius: '50%',
        backgroundColor: 'rgba(255, 255, 255, 0.12)',
        backdropFilter: 'blur(10px)',
        WebkitBackdropFilter: 'blur(10px)',
        border: '1.5px solid rgba(255, 255, 255, 0.25)',
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        zIndex: 2,
        boxShadow: '0 8px 32px 0 rgba(0, 0, 0, 0.15)',
        textAlign: 'center',
        padding: '6px',
        boxSizing: 'border-box',
        transition: 'transform 0.4s cubic-bezier(0.25, 0.8, 0.25, 1)',
        transform: hoveredNode !== null ? 'scale(1.05)' : 'scale(1)',
      }}>
        <div style={{
          fontSize: '0.68rem',
          fontWeight: 700,
          color: '#ffffff',
          letterSpacing: '0.3px',
          textTransform: 'uppercase',
          lineHeight: '1.2',
        }}>
          IT Asset
        </div>
        <div style={{
          fontSize: '0.58rem',
          color: 'rgba(255, 255, 255, 0.8)',
          marginTop: '2px',
          lineHeight: '1.2',
          fontWeight: 500,
        }}>
          Lifecycle
        </div>
      </div>

      {/* Outer Nodes */}
      {nodes.map((node) => {
        const isHovered = hoveredNode === node.id;
        
        return (
          <div
            key={node.id}
            onMouseEnter={() => setHoveredNode(node.id)}
            onMouseLeave={() => setHoveredNode(null)}
            style={{
              position: 'absolute',
              left: `${node.left}px`,
              top: `${node.top}px`,
              transform: `translate(-50%, -50%) ${isHovered ? 'scale(1.1)' : 'scale(1)'}`,
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              width: '80px',
              cursor: 'pointer',
              zIndex: 3,
              transition: 'all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1)',
            }}
          >
            {/* Circular Icon Container */}
            <div style={{
              width: '36px',
              height: '36px',
              borderRadius: '50%',
              backgroundColor: isHovered ? node.color : 'rgba(255, 255, 255, 0.14)',
              border: `1.5px solid ${isHovered ? '#ffffff' : 'rgba(255, 255, 255, 0.3)'}`,
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              boxShadow: isHovered 
                ? `0 0 14px ${node.color}, 0 4px 8px rgba(0,0,0,0.15)` 
                : '0 4px 6px rgba(0,0,0,0.05)',
              transition: 'all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1)',
              color: isHovered ? '#1e293b' : '#ffffff',
            }}>
              <Icon 
                iconName={node.icon} 
                style={{ 
                  fontSize: '15px',
                  fontWeight: isHovered ? 'bold' : 'normal',
                }} 
              />
            </div>

            {/* Label */}
            <span style={{
              marginTop: '5px',
              fontSize: '0.56rem',
              fontWeight: isHovered ? 700 : 500,
              color: isHovered ? '#ffffff' : 'rgba(255, 255, 255, 0.85)',
              textAlign: 'center',
              lineHeight: '1.2',
              textShadow: '0 1px 2px rgba(0,0,0,0.2)',
              transition: 'all 0.2s ease',
              display: 'block',
              maxWidth: '78px',
            }}>
              {node.label}
            </span>
          </div>
        );
      })}
    </div>
  );
};
